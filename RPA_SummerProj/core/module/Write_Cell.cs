using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using Excel = Microsoft.Office.Interop.Excel;
namespace RPA_SummerProj.core.module
{

    public sealed class Write_Cell : CodeActivity
    {
        public InArgument<object> instance { get; set; }

        public InArgument<string> instanceName { get; set; }
        public InArgument<string> SheetName { get; set; }
        public InArgument<string> Cell { get; set; }
        public InArgument<object> Data { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var EngineInstance = (Program)instance.Get(context);
            string InstanceName = instanceName.Get(context);
            string sheetName = SheetName.Get(context);
            string cell = Cell.Get(context);
            object data = Data.Get(context);
            object excel;
            if (EngineInstance.appInstance.TryGetValue(InstanceName, out excel))
            {
                Excel.Application eXL = (Excel.Application)excel;
                Excel.Workbook eWB = eXL.ActiveWorkbook;
                Excel.Worksheet eWS = eWB.Worksheets.Item[sheetName];
                Excel.Range eRng = eWS.Range[cell];
                eRng.Value = data;
            }
            else
                Console.WriteLine("Write Cell Failed");
        }
    }
}
