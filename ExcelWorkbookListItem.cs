using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
namespace ExcelProtocolHelper
{
    class ExcelWorkbookListItem
    {
        public ExcelWorkbook Workbook { get;  }
        public string AlternativeText { get; }

        public ExcelWorkbookListItem(ExcelWorkbook workbook, string alternativeText)
        {
            this.Workbook = workbook;
            this.AlternativeText = alternativeText;
        }

        public override string ToString()
        {
            return Workbook is null ? AlternativeText : Workbook.Name;            
        }

    }
}
