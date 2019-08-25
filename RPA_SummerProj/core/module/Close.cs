using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace RPA_SummerProj.core.module
{

    public sealed class Close : CodeActivity
    {
        public InArgument<object> instance { get; set; }
        public InArgument<string> instanceName { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var EngineInstance = (Program)instance.Get(context);
            string InstanceName = instanceName.Get(context);
            object excel;
            if (EngineInstance.appInstance.TryGetValue(InstanceName, out excel))
            {
                Excel.Application eXL = (Excel.Application)excel;
                Excel.Workbook eWB = eXL.ActiveWorkbook;
                eXL.Quit();
                ReleaseExcelObject(eWB.Worksheets);
                ReleaseExcelObject(eWB);
                ReleaseExcelObject(eXL);
                //ReleaseExcelObject(eWB.Parent);
                EngineInstance.appInstance.Remove(InstanceName);
            }
        }
        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
