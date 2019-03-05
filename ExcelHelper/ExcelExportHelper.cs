using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public static class ExcelExportHelper<T>
    {
        /// <summary>
        /// This is a MIME type of files *.xlsx.
        /// It determines how browsers will process documents.
        /// </summary>
        public static string ExcelContentType
        {
            get
            {
                return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }
        }

        /// <summary>
        /// Returns Excel file as byte array. Mostly for web apps purpose.
        /// </summary>
        public static byte[] ExportToExcel(IEnumerable<T> collection, List<string> columnNames, string worksheetName = "Arkusz1")
        {
            return PrepareExcelFile(worksheetName, collection, columnNames).GetAsByteArray();
        }

        /// <summary>
        /// Saves Excel File into specific directory.
        /// </summary>
        public static void ExportToExcel(IEnumerable<T> collection, List<string> columnNames, DirectoryInfo outputDirectory, string fileName, string worksheetName = "Arkusz1")
        {
            fileName = FileHelper.FixFileNameExcel(fileName);

            FileInfo file = new FileInfo(outputDirectory.FullName + @"\" + fileName);
            FileHelper.DeleteFileIfExist(file);

            PrepareExcelFile(worksheetName, collection, columnNames).Save();
        }

        private static ExcelPackage PrepareExcelFile(string worksheetName, IEnumerable<T> collection, List<string> columnNames)
        {
            using (ExcelPackage excelFile = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelFile.Workbook.Worksheets.Add(worksheetName);
                worksheet.Cells["A1"].LoadFromCollectionFiltered(collection);

                //set column names
                for (int i = 0; i < columnNames.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = columnNames[i];
                }

                worksheet.Cells.AutoFitColumns(0);
                return excelFile;
            }
        }

    }
}
