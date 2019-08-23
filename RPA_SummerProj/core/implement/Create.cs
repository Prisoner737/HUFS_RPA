using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace RPA_SummerProj.core.implement
{
    class Create
    {
        public static void create()
        {
            Excel.Application eXL = null;
            Excel.Workbook eWB = null;
            Excel.Worksheet eWS = null;
            eXL = new Excel.Application();
            eWB = eXL.Workbooks.Add();
            eWS = eWB.Worksheets.Add();
            eXL.Visible = true;
        }
    }
}
