using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using xl = Microsoft.Office.Interop.Excel;

namespace ReadWriteExcel
{
     class ExcelApiTest
    {
        xl.Application xlApp = null;
        xl.Workbook workbook = null;
        xl.Workbooks workbooks = null;
        Hashtable sheets;
        public string xlFilePath;

        public ExcelApiTest(string xlFilePath)
            {
             this.xlFilePath = xlFilePath;
            }

        public void openExcel()
        {
            xlApp = new xl.Application();
            workbooks =xlApp.Workbooks;
            workbook = workbooks.Open(xlFilePath);
            sheets=new Hashtable();
            int count = 1;
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                sheets[count]=sheet.Name;
                count++;
            }
        }
        public void closeExcel() 
        {
            workbook.Close(false,xlFilePath,null);
            Marshal.FinalReleaseComObject(xlApp);
           // workbook = null;


            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            //workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
           // xlApp = null;
        }

        public int getRowCount(string sheetName)
        {
            openExcel();

            int rowCount = 0;
            int sheetValue = 0;

            if(sheets.ContainsValue(sheetName))
            {
                foreach(DictionaryEntry sheet in sheets)
                {
                    if(sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;


                    }
                }
                xl.Worksheet worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;
                rowCount = range.Rows.Count;

            }
            closeExcel();
            return rowCount;
        }
    }
}
