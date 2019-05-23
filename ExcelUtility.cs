using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Xml;
using System.Xml.Serialization;
using System.Configuration;

namespace PtacDealerExcelToTableService
{
    public static class ExcelUtility
    {
        public static List<PtacDealersModel> retrieveExcelData(string excelPath)
        {
            var modelData = new List<PtacDealersModel>();
            FileInfo filePath = new FileInfo(excelPath);
            {
                using (ExcelPackage package = new ExcelPackage(filePath))
                {
                    ExcelWorksheet workSheet = package.Workbook.Worksheets["Sheet1"]; //excel sheet name//
                    int totalRows = workSheet.Dimension.Rows;
                    PtacDealersModel model = new PtacDealersModel();
                    for (int i = 2; i <= totalRows; i++)
                    {
                        modelData.Add(new PtacDealersModel()
                        {
                            DealerName = (workSheet.Cells[i, 2].Value != null) ? ConvertToPascalCase(workSheet.Cells[i, 2].Value.ToString().Trim()) : "",
                            Address = (workSheet.Cells[i, 3].Value != null) ? workSheet.Cells[i, 3].Value.ToString().Trim() : "",
                            City = (workSheet.Cells[i, 4].Value != null) ? workSheet.Cells[i, 4].Value.ToString().Trim() : "",
                            State = (workSheet.Cells[i, 5].Value != null) ? workSheet.Cells[i, 5].Value.ToString().Trim() : "",
                            ZipCode = (workSheet.Cells[i, 6].Value != null) ? workSheet.Cells[i, 6].Value.ToString().Trim() : "",
                            PhoneNumber = (workSheet.Cells[i, 7].Value != null) ? GetPhoneNumber(workSheet.Cells[i, 7].Value.ToString().Trim()) : "",
                            PhoneNumber2 = (workSheet.Cells[i, 8].Value != null) ? GetPhoneNumber(workSheet.Cells[i, 8].Value.ToString().Trim()) : "",

                        });
                    }
                }

            }
            return modelData;
        }

        //method to convert DealerName to Pascal Case & format DealerName as per requirement//
        public static string ConvertToPascalCase(string DealerName)
        {
            string dealerNameLowerCase;
            string[] exceptionalCases = new string[] { "PTAC", "SVC", "SVCS", "LLC", "INC", "Corp", "A/C", "ACR", "CO" };

            // Make DealerName string all lowercase, because ToTitleCase does not change all uppercase correctly //
            dealerNameLowerCase = DealerName.ToLower();
            TextInfo myTextInfo = new CultureInfo("en-US", false).TextInfo;
            dealerNameLowerCase = myTextInfo.ToTitleCase(dealerNameLowerCase);
            var data = GetFormattedDealerName(exceptionalCases, dealerNameLowerCase);

            return data;
        }

        //to format the dealer phone-number//
        public static string GetFormattedDealerName(string[] exceptionalCases, string dealerNameLowerCase)
        {
            string formattedString = dealerNameLowerCase;
            var dealerNames = dealerNameLowerCase.Split(' ');
            foreach (var str in dealerNames)
            {
                bool exists = exceptionalCases.Any(s => s.ToLower().Equals(str.ToLower()));
                if (exists)
                {
                    var tempStr = str.ToUpper();
                    formattedString = formattedString.Replace(str, tempStr);
                }
            }
            return formattedString;
        }

        //method to format Phonenumber to required format//
        public static string GetPhoneNumber(string number)
        {
            string formattedNumber;
            string result = Regex.Replace(number, "[^0-9a-zA-Z]+", "");
            if (string.IsNullOrEmpty(result))
            {
                return number = "";
            }
            else if (result != "NULL" && result.Length >= 10)
            {
                formattedNumber = "(" + result.Substring(0, 3) + ")" + " " + result.Substring(3, 3) + " " + "-" + " " + result.Substring(6, 4);
            }
            else
            {
                return number = "";
            }
            return formattedNumber;
        }


        public static string ToXML<T>(this  List<T> resultantData) where T : class
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlSerializer xmlSerializer = new XmlSerializer(resultantData.GetType());
            using (MemoryStream xmlStream = new MemoryStream())
            {
                xmlSerializer.Serialize(xmlStream, resultantData);
                xmlStream.Position = 0;
                xmlDoc.Load(xmlStream);
                return xmlDoc.InnerXml;
            }
        }
    }
}
