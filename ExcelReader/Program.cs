using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            string path = AppDomain.CurrentDomain.BaseDirectory;
            Console.WriteLine(path);
            XSSFWorkbook xssfwb;

            using (FileStream file = new FileStream(@"TestExcelReader.xlsx", FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = xssfwb.GetSheet("Sheet1");
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {     
                for (int col = 0; col < sheet.GetRow(row).LastCellNum; col++)
                {
                    if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                    {
                        if (sheet.GetRow(row).GetCell(col).CellType.Equals(CellType.String))
                        {
                            Console.Write(string.Format("{0}\t", sheet.GetRow(row).GetCell(col).StringCellValue));
                        }
                        else
                        {
                            Console.Write(string.Format("{0}\t", sheet.GetRow(row).GetCell(col).NumericCellValue));
                        } 
                    }
                }

                Console.WriteLine();
            }


            /*
            var fileName = @"/Users/samiran/Projects/ExcelReader/ExcelReader/TestExcelReader.xlsx";
            var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM [" + sheets.Rows[0]["TABLE_NAME"].ToString() + "] ";

                    var adapter = new OleDbDataAdapter(cmd);
                    var ds = new DataSet();
                    adapter.Fill(ds);
                }
            }
            */
        }
    }
}
