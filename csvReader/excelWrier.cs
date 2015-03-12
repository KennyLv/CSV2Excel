using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Data;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace csvReader
{
    public class excelWrier
    {
        private object missing = Missing.Value;
        private Microsoft.Office.Interop.Excel.Application ExcelRS;
        private Microsoft.Office.Interop.Excel.Workbook RSbook;
        private Microsoft.Office.Interop.Excel.Worksheet RSsheet;

        public excelWrier() { }

        public void doExport(System.Data.DataTable dt, string strExcelFilePath)
        {
            ExcelRS = new Microsoft.Office.Interop.Excel.Application();
            int rowIndex = 1;
            int colIndex = 0;

            RSbook = ExcelRS.Application.Workbooks.Add(missing);
            //RSbook = ExcelRS.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory+@"\TEMP.xlsx", missing, missing, missing, missing, missing,
            //        missing, missing, missing, missing, missing, missing, missing, missing, missing);
           
            RSsheet = (Microsoft.Office.Interop.Excel.Worksheet)RSbook.Sheets.get_Item(1);
            RSsheet.Activate();
            try
            {
                foreach (DataColumn col in dt.Columns)
                {
                    colIndex++;
                    RSsheet.Cells[1, colIndex] = col.ColumnName;
                }

                foreach (DataRow row in dt.Rows)
                {
                    rowIndex++;
                    colIndex = 0;
                    foreach (DataColumn col in dt.Columns)
                    {
                        colIndex++;
                        RSsheet.Cells[rowIndex, colIndex] = row[col.ColumnName].ToString();
                    }
                }
                RSbook.SaveAs(strExcelFilePath, missing, missing, missing, missing, missing, XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
                RSbook.Save();
                RSbook.Close(false, missing, missing);
                ExcelRS.Workbooks.Close();
                ExcelRS.Quit();
            }
            catch
            {

            }
            finally
            {
                //relaese the resource
                RSsheet = null;
                RSbook = null;
                ExcelRS = null;
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(Global.xlSheet);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(Global.xlBook);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(Global.xlBooks);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(Global.xlApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
