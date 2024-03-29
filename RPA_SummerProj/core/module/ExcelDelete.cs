﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using Excel = Microsoft.Office.Interop.Excel;
namespace RPA_SummerProj.core.module
{

    public sealed class ExcelDelete : CodeActivity
    {
        public InArgument<object> instance { get; set; }
        public InArgument<string> instanceName { get; set; }
        public InArgument<string> SheetName { get; set; }
        protected override void Execute(CodeActivityContext context)
        {

            Console.WriteLine("Delete");
            var EngineInstance = (Program)instance.Get(context);
            string InstanceName = instanceName.Get(context);
            string sheetName = SheetName.Get(context);
            object excel;
            //Console.WriteLine(sendingInstance.GetType());

            if (EngineInstance.appInstance.TryGetValue(InstanceName, out excel))
            {
                Console.WriteLine("Success");
                //Excel.Workbook eWB = (Excel.Workbook)excel;
                Excel.Application eXL = (Excel.Application)excel;
                Excel.Workbook eWB = eXL.ActiveWorkbook;
                Excel.Worksheet eWS = eWB.Worksheets.Item[sheetName];
                eWS.Delete();
                //ReleaseExcelObject(eWB);
            }
            else
            {
                Console.WriteLine("Fail");
            }
        }
    }
}
