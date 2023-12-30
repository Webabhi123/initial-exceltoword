using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;
using WordprocessingText = DocumentFormat.OpenXml.Wordprocessing.Text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.Model;
using System.Linq.Expressions;

namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        private System.Windows.Forms.RichTextBox ResultsRichTextBox;
        public Form1()
        {
            InitializeComponent();
            ResultsRichTextBox = new RichTextBox();

        }
        
        private void ProcessButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Open Excel file dialog
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                    openFileDialog.Title = "Select an Excel File";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string excelFilePath = openFileDialog.FileName;

                        // Read Excel data
                        List<Exceltoword.ExcelRowData> excelData = ReadExcel(excelFilePath);

                        // Open Word file dialog
                        using (OpenFileDialog openWordDialog = new OpenFileDialog())
                        {
                            openWordDialog.Filter = "Word Files|*.docx";
                            openWordDialog.Title = "Select a Word Document";

                            if (openWordDialog.ShowDialog() == DialogResult.OK)
                            {
                                string wordTemplatePath = openWordDialog.FileName;

                                // Process each row of Excel data
                                foreach (var excelRow in excelData)
                                {
                                    // Load Word document for each row
                                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordTemplatePath, true))
                                    {
                                        // Access Word document content
                                        Body body = wordDoc.MainDocumentPart.Document.Body;

                                        // Replace placeholders in Word document based on Excel data
                                        ReplacePlaceholderInWordDocument(body,"Name",$"{excelRow.FirstName} {excelRow.LastName}");
                                        ReplacePlaceholderInWordDocument(body,"Contact",excelRow.PhoneNumber);
                                        ReplacePlaceholderInWordDocument(body, "Village", excelRow.Village);
                                        // Add more replacements as needed

                                        // Save modified Word document with a unique name
                                        string outputWordPath = $"{excelRow.FirstName}_{excelRow.LastName}_{excelRow.PhoneNumber}_Modified.docx";
                                        wordDoc.Save();
                                    }
                                }

                                // Display completion message
                                ResultsRichTextBox.Text = $"Modified Word Documents saved at the respective paths.";
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

            private List<Exceltoword.ExcelRowData> ReadExcel(string filePath)
        {
            List<Exceltoword.ExcelRowData> data = new List<Exceltoword.ExcelRowData>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var cells = worksheetPart.Worksheet.Descendants<Cell>();

                //Exceltoword.ExcelRowData currentRow = new Exceltoword.ExcelRowData();
                Dictionary<string, int> columnIndices = new Dictionary<string, int>();

                
                // Get column indices based on the first row
                foreach (Cell cell in cells.Where(c => c.CellReference != null))
                {
                    string cellReference = cell.CellReference.InnerText;
                    string columnName = GetColumnName(cellReference);
                    if (!columnIndices.ContainsKey(columnName))
                    {
                        columnIndices.Add(columnName, Int32.Parse(cellReference.Substring(columnName.Length)));
                    }
                }

                Exceltoword.ExcelRowData currentRow = new Exceltoword.ExcelRowData();


                foreach (Cell cell in cells)
                {
                    string cellValue = GetCellValue(cell, workbookPart);

                    // Check if the cell is in the target row
                    int rowNumber = Int32.Parse(cell.CellReference.InnerText
                .Where(char.IsDigit)
                .Aggregate("", (c, d) => c + d));
                    string columnName = GetColumnName(cell.CellReference.InnerText);


                    if (rowNumber == 1)
                    {
                        // Check specific columns for FirstName, LastName, PhoneNumber, and Village
                        switch (columnName)
                        {
                            case "C":
                                currentRow.FirstName = cellValue;
                                break;
                            case "D":
                                currentRow.LastName = cellValue;
                                break;
                            case "E":
                                currentRow.PhoneNumber = cellValue;
                                break;
                            case "M":
                                currentRow.Village = cellValue;
                                break;
                                // Add more cases for other columns as needed
                        }
                    }

                    // Add more conditions for other columns as needed

                    // Assume a new row after processing a certain number of columns
                    if (cell.CellReference.Value.EndsWith("2"))
                    {
                        data.Add(currentRow);
                        currentRow = new Exceltoword.ExcelRowData();
                    }
                }
            }

            return data;
        }
        private string GetColumnName(string cellReference)
        {
            return new String(cellReference.Where(Char.IsLetter).ToArray());
        }
        private void ReplacePlaceholderInWordDocument(Body body, string placeholder,string value)
        {
        foreach (var textElement in body.Descendants<WordprocessingText>())
        {
            if (textElement.Text.Contains(placeholder))
            {
                    textElement.Text=textElement.Text.Replace(placeholder, value);
                }
        }
    }
        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
            string cellValue = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int sharedStringIndex;
                if (int.TryParse(cellValue, out sharedStringIndex))
                {
                    cellValue = stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
                }
            }

            return cellValue;
        }
    }

}

