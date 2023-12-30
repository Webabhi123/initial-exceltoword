using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Model
{
    class Exceltoword
    {
        public class ExcelRowData
        {
            public string FirstName { get; set; }
            public string  LastName { get; set; }
            public string PhoneNumber { get; set; }

            public string Village { get; set; }

            // Add more properties for other columns as needed
            // Example: public string Column2 { get; set; }
        }
    }
}
