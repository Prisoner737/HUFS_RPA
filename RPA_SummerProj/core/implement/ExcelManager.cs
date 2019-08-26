using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace RPA_SummerProj.core.implement
{
    class ExcelManager
    {
        static Excel.Application eXL = null;
        static Excel.Workbook eWB = null;
        static Excel.Worksheet eWS = null;
        static Excel.Range eRng;

        #region Primitive Function
        //새로운 excel worksheet하나생성
        public static void Create()
        {
            eXL = new Excel.Application();
            eWB = eXL.Workbooks.Add();
            eWS = eWB.Worksheets.Add();
            eXL.Visible = true;
        }

        //기존 excel file open
        public static void Open(string path, string sheetName = null)
        {
            eXL = new Excel.Application();
            eWB = eXL.Workbooks.Open(path);
            eWS = null;
            eXL.Visible = true;
            if (sheetName == null)  //default : 현재 workbook의 첫번째 worksheet를 open
                eWS = eWB.Worksheets.get_Item(1) as Excel.Worksheet;
            else                    //workbook 내에 여러 시트중 원하는 시트가 있으면 해당 시트 open
                eWS = eWB.Worksheets.Item[sheetName];
        }

        //현재 Active한 excel file에 새로운 worksheet append
        public static void Append()
        {
            Excel.Worksheet newWS = null;
            newWS = eWB.Worksheets.Item[eWB.Worksheets.Count];
            eWS = eWB.Worksheets.Add();
            eWS.Move(After: newWS);
        }

        //Delete specific worksheet
        public static void Delete(string sheetName)
        {
            eWS = eWB.Worksheets.Item[sheetName];
            eWS.Delete();
        }

        //Save WorkBook
        public static void Save(string path)
        {
            if (path == null)
            {
                eWB.Save();
            }
            else
            {
                eWB.SaveAs(path);
            }
        }
        //Close WorkBook
        public static void Close()
        {
            eWB.Close();
            ReleaseExcelObject(eWS);
            ReleaseExcelObject(eWB);
            ReleaseExcelObject(eXL);
        }
        //Release Excel Process Completely
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

        #endregion

        #region Range


        public static void Append_Range(string sheetName, string start, string end)
        {
            eWS = eWB.Worksheets.Item[sheetName];
            eRng = eXL.Union(eRng, eWS.get_Range(start, end));
        }
        #endregion

        #region Modify Cell
        public static object Read_Cell(string sheetName, string cell)
        {
            eWS = eWB.Worksheets.Item[sheetName];
            eRng = eWS.Range[cell];
            return eRng.Value;
        }
        public static object[,] Read_Range(string sheetName, string start, string end)
        {
            eWS = eWB.Worksheets.Item[sheetName];
            eRng = eWS.get_Range(start, end);
            return eRng.Value;
        }
        public static void Write_Cell(string sheetName, string cell, object data)
        {
            eWS = eWB.Worksheets.Item[sheetName];
            eRng = eWS.Range[cell];
            eRng.Value = data;
        }

        public static void Write_Range(string sheetName, string start, string end, object[,] data)
        {
            eWS = eWB.Worksheets.Item[sheetName];
            eRng = eWS.get_Range(start, end);
            eRng.Value = data;
        }

        public static void Range_Char_Color(string sheetName, string start, string end, Color color)
        {
            eWS = eWB.Worksheets.Item[sheetName];
            eRng = eWS.get_Range(start, end);
            eRng.Characters.Font.Color = color;
        }

        public static void Range_Cell_Color(string sheetName, string start, string end, Color color)
        {
            eWS = eWB.Worksheets.Item[sheetName];
            eRng = eWS.get_Range(start, end);
            eRng.Interior.Color = color;
        }

        #endregion
    }
}
