using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Validator
{
    class Program
    {

        static void Main(string[] args)
        {
            if (args.Length != 4)
            {
                Console.WriteLine("Usage: Validator.exe \"targetYear\" \"targetMonth\" \"inputFileName\" \"templateFileName\"");
                return;
            }

            var targetYear = int.Parse(args[0]); // 1st argument
            var targetMonth = int.Parse(args[1]); // 2nd argument
            int ok = 0, wrongDate = 0, undefim = 0, duplicate = 0, lackval = 0;
            FileInfo input = new FileInfo(args[2]); // 3rd argument
            FileInfo template = new FileInfo(args[3]); // 4th argument
            Dictionary<string, string> map = new Dictionary<string, string>();

            //template
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();
                var start = excelWorksheet.Dimension.Start;
                var end = excelWorksheet.Dimension.End;
                for (int row = start.Row + 1; row <= end.Row; row++)
                { // Row by row
                    var key = PrepareKey(excelWorksheet, row, 1, 14);
                    if (map.ContainsKey(key))
                        duplicate++;
                    else
                        map.Add(key, string.Empty);
                }
            }

            //input
            using (ExcelPackage excelPackage = new ExcelPackage(input))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();
                var start = excelWorksheet.Dimension.Start;
                var end = excelWorksheet.Dimension.End;
                for (int row = start.Row + 1; row <= end.Row; row++)
                { // Row by row
                    var year = int.Parse(excelWorksheet.Cells[row, 15].Text);
                    var month = int.Parse(excelWorksheet.Cells[row, 16].Text);
                    if (year != targetYear || month != targetMonth)
                    { //check date
                        wrongDate++;
                        continue;
                    }

                    var value = excelWorksheet.Cells[row, 17].Text;
                    var key = PrepareKey(excelWorksheet, row, 1, 14);
                    if (map.ContainsKey(key) && !string.IsNullOrEmpty(map[key])) //key exists, value not empty
                    {
                        duplicate++;
                        Console.WriteLine("Duplicated row: " + row);
                        continue;
                    }
                    
                    if(!map.ContainsKey(key)) //key does not exist
                    {
                        undefim++;
                        Console.WriteLine("Undefined row: " + row);
                        continue;
                    }

                    map[key] = value; //set value
                }

                lackval = map.Keys.Count(x => string.IsNullOrEmpty(map[x])); // count empty values;
                ok = map.Keys.Count(x => !string.IsNullOrEmpty(map[x])); // count not empty values;
            }

            //summary
            Console.WriteLine(string.Format("Read {0} proper records", ok));
            Console.WriteLine(string.Format("\t{0} records out of report date range", wrongDate));
            Console.WriteLine(string.Format("\t{0} records with undefined attribute values", undefim));
            Console.WriteLine(string.Format("\t{0} duplicate records", duplicate));
            Console.WriteLine(string.Format("\t{0} values lacking", lackval));

            //Console.WriteLine(string.Join("\n", map.Keys.Where(x => string.IsNullOrEmpty(map[x]))));

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();

        }

        public static string PrepareKey(ExcelWorksheet excelWorksheet, int row, int colStart, int colEnd)
        {
            var retVal = string.Empty;
            for (int col = colStart; col <= colEnd; col++)
                retVal += excelWorksheet.Cells[row, col].Text;
            return retVal;
        }
    }
}
