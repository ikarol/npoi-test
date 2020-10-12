using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using NPOI.XSSF.Streaming;
using System;
using System.IO;

namespace dotnet
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook wb = new XSSFWorkbook();
            prepareWb(wb, "XSSFtest");
            IWorkbook swb = new SXSSFWorkbook();
            prepareWb(swb, "SXSSFtest");
        }

        private static void prepareWb(IWorkbook wb, string fName) {
            ISheet sheet = wb.CreateSheet("Sheet 1");
            List<string> headers = new List<string>(){
                "A",
                "B",
                "C",
                "D"
            };
            List<string> data = new List<string>(){
                "1",
                "2",
                "",
                "4"
            };
            insertData(sheet, headers, 0);
            insertData(sheet, data, 1);
            string fileName = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\")) + String.Format(
                "{0}-{1}.xlsx",
                DateTime.Now.ToString("yyyy-MMMM-dd HH-mm-ss"),
                Path.GetFileName(fName)
            );
            using (FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                wb.Write(fs);
                fs.Close();
                wb.Close();
                sheet = null;
                wb = null;
            }
        }
        private static void insertData(ISheet st, List<string> data, int rowNum)
        {
            IRow row = st.CreateRow(rowNum);
            int cellNum = 0;
            foreach (string Value in data) {
                if (Value.Equals("")) {
                    cellNum++;
                    continue;
                }
                ICell cell = row.CreateCell(cellNum);
                cell.SetCellValue(Value);
                cellNum++;
            }
        }
    }
}
