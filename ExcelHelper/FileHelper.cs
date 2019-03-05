using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    static class FileHelper
    {
        public static FileInfo DeleteFileIfExist(FileInfo newFile)
        {
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(newFile.FullName);
            }
            return newFile;
        }

        public static string FixFileNameExcel(string fileName)
        {
            if (Path.HasExtension(fileName))
            {
                string extension = Path.GetExtension(fileName);
                if (extension != ".xlsx")
                    fileName = Path.GetFileNameWithoutExtension(fileName) + ".xlsx";
            }
            else
                fileName += ".xlsx";

            return fileName;
        }
    }
}
