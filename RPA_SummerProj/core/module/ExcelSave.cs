using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using Excel = Microsoft.Office.Interop.Excel;
namespace RPA_SummerProj.core.module
{

    public sealed class ExcelSave : CodeActivity
    {
        public InArgument<object> instance { get; set; }
        public InArgument<string> Path { get; set; }
        public InArgument<string> instanceName { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            Console.WriteLine("Save");
            var EngineInstance = (Program)instance.Get(context);
            string path = Path.Get(context);
            string InstanceName = instanceName.Get(context);
            object excel;
            if (EngineInstance.appInstance.TryGetValue(InstanceName, out excel))
            {
                Console.WriteLine("Success");
                Excel.Application eXL = (Excel.Application)excel;
                Excel.Workbook eWB = eXL.ActiveWorkbook;
                if (path == null)
                    eWB.Save();
                else
                    eWB.SaveAs(path);
            }
            else
                Console.WriteLine("Save Failed");
        }
    }
}
