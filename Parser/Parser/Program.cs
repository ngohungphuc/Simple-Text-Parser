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
            CreateExcelFile();
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
            string[] lines = System.IO.File.ReadAllLines(@"D:\SourceCode\Parser\test.txt");

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
        }

        public void ReadAllFile()
        {
            foreach (string fileName in Directory.GetFiles("D:\\SourceCode\\Simple-Text-Parser\\Parser"))
            {
                string[] fileLines = File.ReadAllLines(fileName);
                Console.WriteLine(fileLines);
                // Do something with the file content
            }
        }

        public static void CreateExcelFile()
        {
            //([Date: ])([A - za - z0 - 9 +])\w +
            string file = "D:\\SourceCode\\Simple-Text-Parser\\Parser\\File\\newdoc.xls";
            string[] lines = File.ReadAllLines(@"D:\SourceCode\Simple-Text-Parser\Parser\File\LEADS\[Drip] 4ebdecc7@opayq.com.eml");
            foreach (var line in lines)
            {
                MatchCollection matchCollection = Regex.Matches(line, "(Date|city|country|first_name|last_name|last_trading_or_investment_book_read|state|street1|zip_code|id""):( )([ A-Za-z0-9\\+\\:\\,]+)", RegexOptions.Multiline);
                foreach (var result in matchCollection)
                {
                    Console.WriteLine(result);
                }
            }
            //([A - za - z0 - 9\+\,\:\@\.\"\-]+)?
            // ^ (Date):( )([A - Za - z0 - 9\+\:\,] +)
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
    }
}