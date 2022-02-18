using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;

namespace Automation.Ui.Accelerators.UtilityClasses
{
    public class FileReaderWriter
    {

        public List<string> ReadAllLinesInTextFile(string textFilepath)
        {
            if (string.IsNullOrEmpty(textFilepath))
            {
                throw new ArgumentNullException("Argument textFilepath cannot be empty or null");
            }

            return File.ReadAllLines(textFilepath).ToList();
        }

        public bool CopyFileToCurrentWorkingDirectory(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("Argument filePath cannot be empty or null");
            }

            var fileInfo = new FileInfo(filePath);
            string currentDirPath = Directory.GetCurrentDirectory() + "\\" + fileInfo.Name;
            if (!File.Exists(filePath))
            {
                throw new Exception($"File does not exit {filePath}");
            }

            File.Delete(currentDirPath);

            File.Copy(filePath, currentDirPath);

            return File.Exists(currentDirPath);
        }

        public void DeleteFiles(string dirPath, string[] fileNameList)
        {
            if (string.IsNullOrEmpty(dirPath))
            {
                throw new Exception("Path is null");
            }

            if (fileNameList.Length <= 0)
            {
                throw new Exception("File list is empty");
            }

            foreach (string fileName in fileNameList)
            {
                var filePath = $"dirPath \\{0} fileName";
                var fileInfo = new FileInfo(filePath);

                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
        }

        public void DeleteFiles(string dirPath, string pattern)
        {
            if (string.IsNullOrEmpty(dirPath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("dirpath is empty");
                Console.ResetColor();
                throw new Exception($"Dir path is empty");
            }
            if (!Directory.Exists(dirPath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("dirpath do not exist\n");
                Console.WriteLine("Creating one");
                Directory.CreateDirectory(dirPath);
                Console.ResetColor();
            }

            var fileNameList = Directory.GetFiles(dirPath, pattern);

            Array.ForEach(fileNameList, f =>
            {
                File.Delete(f);
                Console.WriteLine("{0} got deleted", f);
            });
        }


        public static List<bool> GetImageHash(Bitmap BMPSource)
        {
            var listResult = new List<bool>();

            var bmpMin = new Bitmap(BMPSource, new Size(16, 16));
            for (int j = 0; j < bmpMin.Height; j++)
            {
                for (int i = 0; i < bmpMin.Width; i++)
                {
                    listResult.Add(bmpMin.GetPixel(i, j).GetBrightness() < 0.5f);
                }
            }
            return listResult;
        }

        public static bool CreateAFile(string fileName, string fileContent)
        {
            if (string.IsNullOrEmpty(fileName) && string.IsNullOrEmpty(fileContent))
            {
                throw new Exception("Arguments cannot be null or empty");
            }
            bool result = false;

            using (StreamWriter sr = new StreamWriter(fileName))
            {
                sr.Write(fileContent);
                sr.Flush();
                result = true;
            }

            return result;
        }

        public static string ReadContentFromFile(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                throw new Exception("Arguments cannot be null or empty");
            }

            string content = string.Empty;
            using (StreamReader str = new StreamReader(fileName))
            {
                content = str.ReadToEnd();
            }

            return content;
        }
    }
}