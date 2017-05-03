using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;
using Excel = Microsoft.Office.Interop.Excel;

namespace Parser
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ReadAllFile();
            //CreateExcelFile();
            //ConvertToXml();
        }

        public static void TrimHtmlTag()
        {
            string html = File.ReadAllText(@"D:\SourceCode\Simple-Text-Parser\Parser\2618.eml");
            StringBuilder pureText = new StringBuilder();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);

            foreach (HtmlNode node in doc.DocumentNode.ChildNodes)
            {
                pureText.Append(node.InnerText);
            }
        }

        public static void ConvertToXml()
        {
            string[] lines = System.IO.File.ReadAllLines(@"E:\Source Code\Study\Simple-Text-Parser\Parser\Data\data.txt");

            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xmlDocument.DocumentElement;
            xmlDocument.InsertBefore(xmlDeclaration, root);
            XmlElement bodyElement = xmlDocument.CreateElement(string.Empty, "items", string.Empty);
            xmlDocument.AppendChild(bodyElement);

            var i = 0;

            XmlElement itemElement = null;
            foreach (string line in lines)
            {
                var result = line.Split(new[] { ":" }, StringSplitOptions.None);

                if (result.Length == 1)
                {
                    if (result[0] == string.Empty)
                    {
                        i = 0;
                        continue;
                    }
                    continue;
                }

                if (i == 0)
                {
                    itemElement = xmlDocument.CreateElement(string.Empty, "item", string.Empty);
                    bodyElement.AppendChild(itemElement);
                }

                if (lines[i] != null && result[1] != null)
                {
                    if (itemElement != null)
                    {
                        string elementName = result[0].Replace(" ", "-");
                        string elementValue = result[1];
                        XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(), string.Empty);
                        element.InnerText = elementValue;
                        itemElement.AppendChild(element);
                    }
                }

                i++;
            }
            xmlDocument.Save(@"E:\Source Code\Study\Simple-Text-Parser\Parser\Data\data.xml");
        }

        public static void ReadAllFile()
        {
            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xmlDocument.DocumentElement;
            xmlDocument.InsertBefore(xmlDeclaration, root);
            XmlElement bodyElement = xmlDocument.CreateElement(string.Empty, "items", string.Empty);
            xmlDocument.AppendChild(bodyElement);
            XmlElement itemElement = null;
            var i = 0;

            foreach (string fileName in Directory.GetFiles("E:\\Source Code\\Study\\Simple-Text-Parser\\Parser\\Data\\test"))
            {
                string[] fileLines = File.ReadAllLines(fileName);
                StreamReader sr = new StreamReader(fileName);
                foreach (var line in fileLines)
                {
                    MatchCollection matchCollections = Regex.Matches(line, "(Email|IP Address|city|country|first_name|has_read_LBoT|has_read_TC|has_read_TF|last_name|last_trading_or_investment_book_read|state|street1|street2|zip_code|Date|utc_offset|visitor_uuid|time_zone|friendly_time_zone|tags|created_a)(:)( )([ A-Za-z0-9\\&\\+\\,\\:\\@\\.\\\"\\-]+)?", RegexOptions.Multiline);

                    foreach (var lineResult in matchCollections)
                    {
                        var lineResultParse = lineResult.ToString();

                        if (lineResultParse.Equals(string.Empty)) continue;

                        var result = lineResultParse.Split(new[] { ":" }, StringSplitOptions.None);
                        if (result.Length == 1)
                        {
                            if (result[0] == string.Empty)
                            {
                                i = 0;
                                continue;
                            }
                            continue;
                        }

                        if (i == 0)
                        {
                            itemElement = xmlDocument.CreateElement(string.Empty, "item", string.Empty);
                            bodyElement.AppendChild(itemElement);
                        }

                        if (itemElement != null)
                        {
                            string elementName = result[0].Replace(" ", "-");
                            string elementValue = result[1];
                            XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(), string.Empty);
                            element.InnerText = elementValue;
                            itemElement.AppendChild(element);
                        }

                        i++;

                        //if (sr.EndOfStream)
                        //{
                        //    var fileNameExtract = Path.GetFileNameWithoutExtension(fileName);
                        //    xmlDocument.Save($"E:\\Source Code\\Study\\Simple-Text-Parser\\Parser\\Data\\Result\\{fileNameExtract}.xml");
                        //    i = 0;
                        //}
                    }
                }
                var fileNameExtract = Path.GetFileNameWithoutExtension(fileName);
                xmlDocument.Save($"E:\\Source Code\\Study\\Simple-Text-Parser\\Parser\\Data\\Result\\{fileNameExtract}.xml");
                i = 0;
            }
        }

        public static void CreateExcelFile()
        {
            string file = "D:\\SourceCode\\Simple-Text-Parser\\Parser\\File\\newdoc.xls";
            string[] lines = File.ReadAllLines(@"E:\Source Code\Study\Simple-Text-Parser\Parser\Data\LEADS\data2 (0).eml");
            foreach (var line in lines)
            {
                MatchCollection matchCollection = Regex.Matches(line, "(Email|IP Address|city|country|first_name|has_read_LBoT|has_read_TC|has_read_TF|last_name|last_trading_or_investment_book_read|state|street1|street2|zip_code|Date|utc_offset|visitor_uuid|time_zone|friendly_time_zone|tags|created_a)(:)( )([ A-Za-z0-9\\&\\+\\,\\:\\@\\.\\\"\\-]+)?", RegexOptions.Multiline);
                foreach (var result in matchCollection)
                {
                    Console.WriteLine(result);
                }
            }
            //Excel.Application oApp;
            //Excel.Worksheet oSheet;
            //Excel.Workbook oBook;

            //oApp = new Excel.Application();
            //oBook = oApp.Workbooks.Add();
            //oSheet = (Excel.Worksheet)oBook.Worksheets.Item[1];
            //oSheet.Cells[1, 1] = "some value";
            //oBook.SaveAs(file);
            //oBook.Close();
            //oApp.Quit();
        }

        //(Email|IP Address|city|country|first_name
        //|has_read_LBoT|has_read_TC|has_read_TF|last_name
        //|last_trading_or_investment_book_read|state|street1|street2|zip_code|Date|
        //utc_offset|visitor_uuid|time_zone|friendly_time_zone|tags|created_at)(:)( )([ A-Za-z0-9\\&\\+\\,\\:\\@\\.\\\"\\-]+)?
    }
}