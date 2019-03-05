using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public static class ExcelImportHelper<T> where T : new()
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
        /// Returns list of objects based on *.xlsx worksheet.
        /// </summary>
        public static List<T> ImportFromExcel(string fileName)
        {
            using (FileStream fileStream = new FileStream(fileName, FileMode.Open))
            {
                ExcelPackage excel = new ExcelPackage(fileStream);
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[1];

                return worksheet.ConvertSheetToObjects<T>().ToList();
            }
        }
    }
}
