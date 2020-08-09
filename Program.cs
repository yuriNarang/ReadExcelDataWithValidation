using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter Path of Excel File:");
            string path = Console.ReadLine();
            Console.WriteLine("Enter no. of rows");
            int rows = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter no. of columns");
            int cols = Convert.ToInt32(Console.ReadLine());
            ReadExcelData(path, rows, cols);

            Dal.GetAppSettingsFile();
        }
        
        private static void ReadExcelData(string path, int rows, int cols)
        {
            string errStr = string.Empty;
            try
            {
                if (!File.Exists(path))
                {
                    Console.WriteLine("File not exist.");
                    return;
                }
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(path)))
                {
                    var firstSheet = package.Workbook.Worksheets[0];
                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= cols; j++)
                        {
                            if (j == 1) // column 1
                            {
                                errStr = validateDate(firstSheet.GetValue(j, i).ToString());
                                if (errStr.Length > 0)
                                {
                                    Console.WriteLine(errStr);
                                    return;
                                }
                            }
                            if (j == 2) // column 2
                            {
                                errStr = validateNumber(firstSheet.GetValue(j, i).ToString());
                                if (errStr.Length > 0)
                                {
                                    Console.WriteLine(errStr);
                                    return;
                                }
                            }
                            if (j == 3) // column 3
                            {
                                errStr = validateString(firstSheet.GetValue(j, i).ToString());
                                if (errStr.Length > 0)
                                {
                                    Console.WriteLine(errStr);
                                    return;
                                }
                            }
                        }
                    }
                    Console.WriteLine("");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private static string validateString(string cellValue, bool canContainSpecialCharacter = true, bool canContainNumbers = true, int minLength = 1, int maxLength = int.MaxValue)
        {
            StringBuilder errorLogSB = new StringBuilder();
            string charSet = string.Empty;
            if (!canContainSpecialCharacter)
            {
                Regex r = new Regex(@"[~`!@#$%^&*()-+=|\{}':;.,<>/?]");
                if (r.IsMatch(cellValue))
                {
                    errorLogSB.Append("It contains special character. ");
                }
            }
            if (!canContainNumbers)
            {
                if (cellValue.Any(Char.IsDigit))
                {
                    errorLogSB.Append("It contains numbers. ");
                }
            }
            if (cellValue.Length < minLength)
            {
                errorLogSB.Append("Minimum length should be " + minLength + ". ");
            }
            if (cellValue.Length > maxLength)
            {
                errorLogSB.Append("Maximum length should be " + maxLength + ". ");
            }

            return errorLogSB.ToString();
        }

        private static string validateNumber(string cellValue, bool canContainDecimal = true, int minValue = int.MinValue, int maxValue = int.MaxValue)
        {
            StringBuilder errorLogSB = new StringBuilder();
            if (!canContainDecimal)
            {
                if (cellValue.IndexOf(".") >= 0)
                {
                    errorLogSB.Append("It contains decimal. ");
                }
            }
            if (Convert.ToDouble(cellValue) < minValue)
            {
                errorLogSB.Append("Minimum Value should be " + minValue + ". ");
            }
            if (Convert.ToDouble(cellValue) > maxValue)
            {
                errorLogSB.Append("Maximum value should be " + maxValue + ". ");
            }

            return errorLogSB.ToString();
        }

        private static string validateDate(string cellValue, bool canBeBlankOrNull = false, DateTime? minValue = null, DateTime? maxValue = null)
        {
            StringBuilder errorLogSB = new StringBuilder();
            if (!minValue.HasValue)
            {
                minValue = DateTime.MinValue;
            }
            if (!maxValue.HasValue)
            {
                maxValue = DateTime.MaxValue;
            }
            if (!canBeBlankOrNull)
            {
                if (string.IsNullOrEmpty(cellValue))
                {
                    errorLogSB.Append("Date cannot be null or empty. ");
                }
            }
            if (Convert.ToDateTime(cellValue) < minValue)
            {
                errorLogSB.Append("Minimum Date should be " + minValue + ". ");
            }
            if (Convert.ToDateTime(cellValue) > maxValue)
            {
                errorLogSB.Append("Maximum Date should be " + maxValue + ". ");
            }

            return errorLogSB.ToString();
        }
    }
}
