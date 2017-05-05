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
            LeadsFolder();
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

        public static void LeadsFolder()
        {
            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xmlDocument.DocumentElement;
            xmlDocument.InsertBefore(xmlDeclaration, root);
            XmlElement bodyElement = xmlDocument.CreateElement(string.Empty, "items", string.Empty);
            xmlDocument.AppendChild(bodyElement);
            XmlElement itemElement = null;
            var i = 0;
            try
            {
                foreach (string fileName in Directory.GetFiles("D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\test"))
                {
                    string[] fileLines = File.ReadAllLines(fileName);
                    StreamReader sr = new StreamReader(fileName);

                    foreach (var line in fileLines)
                    {
                        MatchCollection matchCollections = Regex.Matches(line, "(id|email|status|Email|IP Address|IP Address|city|country|first_name|has_read_LBoT|has_read_TC|has_read_TF|last_name|last_trading_or_investment_book_read|state|street1|street2|zip_code|Date|utc_offset|visitor_uuid|time_zone|friendly_time_zone|tags|created_at|lead_score|time_zone)(:)( )?([ A-Za-z0-9\\/\\&\\+\\,\\:\\@\\.\\\"\\-\\&\\[\\]]+)?", RegexOptions.Multiline);

                        foreach (var lineResult in matchCollections)
                        {
                            var lineResultParse = lineResult.ToString();

                            if (lineResultParse.Equals(string.Empty)) continue;

                            var result = lineResultParse.Split(new[] { ":" }, 2, StringSplitOptions.None);
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
                                elementValue = elementValue == string.Empty ? "n/a" : result[1];
                                XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(), string.Empty);
                                element.InnerText = elementValue;
                                itemElement.AppendChild(element);
                            }

                            i++;
                        }
                    }

                    var fileNameExtract = Path.GetFileNameWithoutExtension(fileName);
                    xmlDocument.Save($"D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\Result\\{fileNameExtract}.xml");
                    i = 0;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}