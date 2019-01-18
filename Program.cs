using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;

namespace ExcelToCsv
{
    class Program
    {
        public static List<string> specialFiles = new List<string>()
        { "AttrBase" };

        static void Main(string[] args)
        {
            string filePath;
            if (args != null && args.Length > 0)
            {
                Console.WriteLine(args[0]);
                filePath = args[0].Trim() + "/";
            }
            else
            {
                Console.WriteLine("Input the folder path:");
                filePath = Console.ReadLine().Trim() + "/";
            }
            //string saveFolder = filePath;
            //string fileName = "test.csv";

            DirectoryInfo root = new DirectoryInfo(filePath);

            foreach (FileInfo f in root.GetFiles())
            {
                string name = f.Name;
                if (name.Split('.')[1] == "xlsx")
                {
                    Console.WriteLine("Start converting {0} ...", name);
                    SaveToCsv(GetSheetValues(f), filePath, name.Split('.')[0]);
                    Console.WriteLine("Converted successfully.");
                }
                //Console.WriteLine("name = {0}, fullName = {1}",name,fullName);
            }

            Console.WriteLine("Done");
        }

        public static List<string> GetSheetValues(FileInfo file)
        {
            //FileInfo file = new FileInfo(filepath);
            List<string> sheetString = new List<string>();

            if (file != null)
            {
                bool isSpecialFile = false;
                if (specialFiles.Contains(file.Name.Split('.')[0]))
                {
                    isSpecialFile = true;
                }
                using (ExcelPackage package = new ExcelPackage(file))
                {

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    //获取表格的列数和行数
                    int rowCount = worksheet.Dimension.Rows;
                    int ColCount = worksheet.Dimension.Columns;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        string sheetRowStr = "";
                        //Person person = new Person();
                        for (int col = 1; col <= ColCount; col++)
                        {
                            Object cell = worksheet.Cells[row, col].Value;

                            if (col == 1)
                            {
                                if (cell == null)
                                {
                                    sheetRowStr = "";
                                }
                                else if (cell.ToString() == "-1" && !isSpecialFile)
                                {
                                    sheetRowStr = "";
                                }
                                else
                                {
                                    sheetRowStr = cell.ToString();
                                }
                            }
                            else
                            {
                                sheetRowStr += ",";
                                if (cell == null)
                                {
                                    sheetRowStr += "";
                                }
                                else if (cell.ToString() == "-1" && !isSpecialFile)
                                {
                                    sheetRowStr += "";
                                }
                                else
                                {
                                    sheetRowStr += cell.ToString();
                                }
                            }
                        }
                        sheetString.Add(sheetRowStr);
                    }
                    return sheetString;
                }
            }
            return null;
        }

        public static void SaveToCsv(List<string> sheetString, string saveFolder, string fileName)
        {
            string filePath = saveFolder + fileName + ".csv";

            FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);

            for (int i = 0; i < sheetString.Count; i++)
            {
                if (i == sheetString.Count - 1)
                {
                    sw.Write(sheetString[i]);
                }
                else
                {
                    sw.WriteLine(sheetString[i]);
                }
            }

            sw.Close();
            fs.Close();
        }

    }
}
