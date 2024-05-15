using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulkUploaddata
{
    internal class Config
    {
        public static string getRootFolder()
        {
            String currpath = System.IO.Directory.GetCurrentDirectory();
            return currpath.Split(new string[] { "\\bin" }, StringSplitOptions.None)[0];
        }

        public static bool RenameFile(string filePath, string newFileName)
        {
            bool status = false;
            if (File.Exists(filePath))
            {
                try
                {
                    string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
                    File.Move(filePath, newFilePath);
                    status = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                    status = false;
                }
            }
            return status;
        }
    }
}
