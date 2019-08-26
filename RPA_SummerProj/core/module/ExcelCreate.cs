using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using Excel = Microsoft.Office.Interop.Excel;
namespace RPA_SummerProj.core.module
{

    public sealed class Create : CodeActivity
    {
        public InArgument<object> instance { get; set; }
        public InArgument<string> instanceName { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            Console.WriteLine("Create");
            var EngineInstance = (Program)instance.Get(context);
            string InstanceName = instanceName.Get(context);
            Excel.Application eXL = new Excel.Application();
            Excel.Workbook eWB = eXL.Workbooks.Add();
            eXL.Visible = true;
            //EngineInstance.appInstance.Add("Excel", eWB);
            EngineInstance.appInstance.Add(InstanceName, eXL);
        }
    }
}
