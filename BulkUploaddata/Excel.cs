using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace BulkUploaddata
{
    internal class Excel
    {
        public static DataTable ExcelDataToDataTable(string filePath, string sheetName, bool hasHeader)
        {
            var dt = new DataTable();
            var fi = new FileInfo(filePath);
            if (!fi.Exists)
                throw new Exception("File " + filePath + " Does Not Exists");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var xlPackage = new ExcelPackage(fi);
            var worksheet = xlPackage.Workbook.Worksheets[sheetName];

            dt = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].ToDataTable(c =>
            {
                c.FirstRowIsColumnNames = true;
            });

            return dt;
        }
        public static string GetExcelSheetHeaders(string filePath, string sheetName)
        {
           
            string[] headers = null;
            string o = "";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                int startColumn = worksheet.Dimension.Start.Column;
                int endColumn = worksheet.Dimension.End.Column;
                int startRow = worksheet.Dimension.Start.Row;
                int endRow = startRow; // Assuming only the first row is headers
                headers = new string[endColumn - startColumn + 1];
                for (int col = startColumn; col <= endColumn; col++)
                {
                    headers[col - startColumn] = worksheet.Cells[startRow, col].Text;
                }
                for(int i=0;i<headers.Length;i++)
                {
                    o=o+","+ headers[i];
                }
            }

            return o;
        }

        public static void WriteResultDataToASheet(string sheetname, DataTable table, string schema, string tablename)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = Config.getRootFolder() + "\\" + ConfigurationManager.AppSettings["resultfolder"] + "\\" + tablename + "_latest"+timestamp+".xlsx";
           
            using (ExcelPackage pck = new ExcelPackage(new System.IO.FileInfo(path)))
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(sheetname);
                ws.Cells["A1"].LoadFromDataTable(table, true);
                pck.Save();
            }
        }

    }
}
