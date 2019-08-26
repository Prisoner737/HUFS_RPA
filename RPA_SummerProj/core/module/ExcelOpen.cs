using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using Excel = Microsoft.Office.Interop.Excel;
namespace RPA_SummerProj.core.module
{

    public sealed class ExcelOpen : CodeActivity
    {
        public InArgument<string> Path { get; set; }
        public InArgument<string> SheetName { get; set; }
        public InArgument<object> instance { get; set; }
        public InArgument<string> instanceName { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            string path = Path.Get(context);
            string sheetName = SheetName.Get(context);
            string InstanceName = instanceName.Get(context);
            var EngineInstance = (Program)instance.Get(context);
            Excel.Application eXL = new Excel.Application();
            eXL.Visible = true;
            Excel.Workbook eWB = eXL.Workbooks.Open(path);
            Excel.Worksheet eWS;
            if (sheetName == null)
                eWS = eWB.Worksheets.get_Item(1);
            else
                eWS = eWB.Worksheets.Item[sheetName];
            EngineInstance.appInstance.Add(InstanceName, eXL);

        }
    }
}
