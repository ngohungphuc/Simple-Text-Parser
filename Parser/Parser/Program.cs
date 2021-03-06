﻿using HtmlAgilityPack;
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
            Orders();
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
                foreach (
                    string fileName in Directory.GetFiles("D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\LEADS")
                )
                {
                    string[] fileLines = File.ReadAllLines(fileName);
                    StreamReader sr = new StreamReader(fileName);

                    foreach (var line in fileLines)
                    {
                        MatchCollection matchCollections = Regex.Matches(line,
                            "(id|email|status|Email|IP Address|IP Address|city|country|first_name|has_read_LBoT|has_read_TC|has_read_TF|last_name|last_trading_or_investment_book_read|state|street1|street2|zip_code|Date|utc_offset|visitor_uuid|time_zone|friendly_time_zone|tags|created_at|lead_score|time_zone)(:)( )?([ A-Za-z0-9\\/\\&\\+\\,\\:\\@\\.\\\"\\-\\&\\[\\]]+)?",
                            RegexOptions.Multiline);

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
                                XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                                    string.Empty);
                                element.InnerText = elementValue;
                                itemElement.AppendChild(element);
                            }

                            i++;
                        }
                    }

                    xmlDocument.Save($"D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\Result\\LEADS.xml");
                    i = 0;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void ORDERS_CC_AUTHORIZE()
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
                foreach (
                    string fileName in
                    Directory.GetFiles(
                        "D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\DRIP_PHUC\\ORDERS_CC_AUTHORIZE.NET")
                )
                {
                    string[] fileLines = File.ReadAllLines(fileName);
                    StreamReader sr = new StreamReader(fileName);

                    foreach (var line in fileLines)
                    {
                        MatchCollection matchCollections = Regex.Matches(line,
                            "(Date|Merchant|Date/Time|Invoice|Description|Amount|Payment Method|Transaction Type|Response|Auth Code|Transaction ID|Address Verification|Customer ID|First Name|Last Name|Company|Address|City|State/Province|Zip/Postal Code|Country|Phone|Fax|E-Mail|First Name|Last Name|Company|Address|City|State/Province|Postal Code|Country|Tax|Duty|Freight|Tax Exempt|PO Number)( )?(:)?( )?([ A-Za-z0-9\\/\\&\\+\\,\\:\\@\\.\\\"\\-\\&\\[\\](\\\\)\\-]+)?",
                            RegexOptions.Multiline);
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
                                string elementName = string.Empty;
                                //\s matches whitespace
                                //+ means one or more of the preceding expression
                                //(one or more whitespace characters)
                                //$ means the end of the string
                                if (Regex.IsMatch(result[0], @"\s+$"))
                                {
                                    elementName = result[0].Replace(" ", "");
                                    elementName = elementName.Replace(elementName.Contains("/") ? "/" : " ", "-");
                                    string elementValue = result[1];
                                    elementValue = elementValue == " " ? "n/a" : result[1];
                                    XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                                        string.Empty);
                                    element.InnerText = elementValue;
                                    itemElement.AppendChild(element);
                                }
                                else
                                {
                                    elementName = result[0].Replace(" ", "-");
                                    string elementValue = result[1];
                                    elementValue = elementValue == string.Empty ? "n/a" : result[1];
                                    XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                                        string.Empty);
                                    element.InnerText = elementValue;
                                    itemElement.AppendChild(element);
                                }
                            }

                            i++;
                        }
                    }
                    xmlDocument.Save(
                        $"D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\Result\\ORDERS_CC_AUTHORIZE.xml");
                    i = 0;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void UnknownEmailList()
        {
            string[] lines =
                System.IO.File.ReadAllLines(
                    @"D:\SourceCode\Simple-Text-Parser\Parser\Data\DRIP_PHUC\LEADS_EUDORA\unknown email list");

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
                if (i == 0)
                {
                    itemElement = xmlDocument.CreateElement(string.Empty, "item", string.Empty);
                    bodyElement.AppendChild(itemElement);
                }

                string elementValue = line;
                XmlElement element = xmlDocument.CreateElement(string.Empty, "email", string.Empty);
                element.InnerText = elementValue;
                itemElement.AppendChild(element);

                i++;
            }
            xmlDocument.Save(@"D:\SourceCode\Simple-Text-Parser\Parser\Data\Result\UnknownEmailList.xml");
        }

        public static void LEADS_EUDORA_Folder()
        {
            //((Date:)()([A-Za-z0-9\\:\\,\\-]+))| (\[that\](.*?)\[has\])|((?=:).*(?=subscribers))LEADS_EUDORA
            string[] lines =
                File.ReadAllLines(@"D:\SourceCode\Simple-Text-Parser\Parser\Data\DRIP_PHUC\LEADS_EUDORA\LEADS copy");
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
                foreach (var line in lines)
                {
                    if (line.Contains("</x-charset>"))
                    {
                        i = 0;
                        continue;
                    }
                    MatchCollection matchCollections = Regex.Matches(line,
                        "((Name|Address|Address Cont|City|State|Country|Zip/Postal Code|Email|Phone|Return Phone Call|What can we do to help you|In Person Seminars|Newsletter Add|New Trader|How long trading|Name of seminars attended|Purchased a trading system|What system did you purchase|Markets traded|Newsletter Add|Recent Trading Books Read|Do you own Trend Following|Do you own TurtleTrader|Do you own Market Wizards|Where did you year about trend following|Do you own mutual funds|Where did you find out about the offering|In Person Seminars|What can we do to help you)(:)?( )?([ A-Za-z0-9\\/\\&\\+\\,\\:\\@\\.\\\"\\-\\&\\[\\](\\\\)\\-]+)?)",
                        RegexOptions.Multiline);

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
                            string elementName;
                            string elementValue;
                            //\s matches whitespace
                            //+ means one or more of the preceding expression
                            //(one or more whitespace characters)
                            //$ means the end of the string

                            elementName = result[0].Replace(" ", "-");
                            elementName =
                                elementName.Replace(
                                    (elementName.Contains("/") || elementName.Contains("?")) ? "/" : " ", "-");
                            elementValue = result[1];
                            elementValue = elementValue == " " ? "n/a" : result[1];
                            XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                                string.Empty);
                            element.InnerText = elementValue;
                            itemElement.AppendChild(element);
                        }

                        i++;
                    }
                }

                xmlDocument.Save($"D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\test\\LEADS copy.xml");
                i = 0;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void TurtleSales()
        {
            string[] lines = File.ReadAllLines(@"D:\SourceCode\Simple-Text-Parser\Parser\Data\test\Turtle Sales.mbox");

            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xmlDocument.DocumentElement;
            xmlDocument.InsertBefore(xmlDeclaration, root);
            XmlElement bodyElement = xmlDocument.CreateElement(string.Empty, "items", string.Empty);
            xmlDocument.AppendChild(bodyElement);

            var i = 0;

            XmlElement itemElement = null;
            try
            {
                foreach (var line in lines)
                {
                    if (line.Contains("From -"))
                    {
                        i = 0;
                        continue;
                    }
                    MatchCollection matchCollections = Regex.Matches(line,
                        "(From - |Date|Product|E-Mail|How do you know Turtle Trader|Phone|First Name|Last Name|Company|Address|City|State|Zip|Payment Method|Account Number|Card Holder|Expiration)(:)( )?([ A-za-z0-9\\+\\,\\:\\@\\.\\\"\\-\\\\;\\\\/]+)?",
                        RegexOptions.Multiline);

                    foreach (var lineResult in matchCollections)
                    {
                        var lineResultParse = lineResult.ToString();

                        if (lineResultParse.Equals(string.Empty)) continue;
                        if (lineResultParse.Contains("From -")) i = 0;

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
                            XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                                string.Empty);
                            element.InnerText = elementValue;
                            itemElement.AppendChild(element);
                        }

                        i++;
                    }
                }

                xmlDocument.Save($"D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\test\\Turtle Sales mbox.xml");
                i = 0;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void NewLetterSignUp()
        {
            //((Date:)()([A-Za-z0-9\\:\\,\\-]+))| (\[that\](.*?)\[has\])|((?=:).*(?=subscribers))
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
                foreach (string fileName in Directory.GetFiles("D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\test")
                )
                {
                    string[] fileLines = File.ReadAllLines(fileName);
                    StreamReader sr = new StreamReader(fileName);

                    foreach (var line in fileLines)
                    {
                        if (line.Contains("-TurtleTrader News"))
                        {
                            i = 0;
                            continue;
                        }
                        MatchCollection matchCollections = Regex.Matches(line,
                            "((Date|Reply-To)(:)?( )?([ A-Za-z0-9\\/\\&\\+\\,\\:\\@\\.\\\"\\-\\&\\[\\](\\\\)\\-]+)?)|((?=:).*(?=subscribers))",
                            RegexOptions.Multiline);

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
                                string elementName;
                                if (result[0] == string.Empty && result[1] != string.Empty)
                                {
                                    elementName = "subscribers";
                                }
                                else
                                {
                                    elementName = result[0].Replace(" ", "-") == "Reply-To"
                                        ? "subscribed"
                                        : result[0].Replace(" ", "-");
                                }

                                string elementValue = result[1];
                                elementValue = elementValue == string.Empty ? "n/a" : result[1];
                                XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                                    string.Empty);
                                element.InnerText = elementValue;
                                itemElement.AppendChild(element);
                            }

                            i++;
                        }
                    }

                    xmlDocument.Save($"D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\test\\test.xml");
                    i = 0;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void UsFilmOrder()
        {
            string[] lines = File.ReadAllLines(@"D:\SourceCode\Simple-Text-Parser\Parser\Data\test\US Film Order");

            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xmlDocument.DocumentElement;
            xmlDocument.InsertBefore(xmlDeclaration, root);
            XmlElement bodyElement = xmlDocument.CreateElement(string.Empty, "items", string.Empty);
            xmlDocument.AppendChild(bodyElement);

            var i = 0;

            XmlElement itemElement = null;

            foreach (var line in lines)
            {
                if (line.Contains("From -"))
                {
                    i = 0;
                    continue;
                }
                MatchCollection matchCollections = Regex.Matches(line,
                    "(Name|Date|Address|Address Cont|City|State|Zip/Postal Code|Country|Phone|Email|Quantity|Amount)(:)( )?([ A-za-z0-9\\+\\,\\:\\@\\.\\\"\\-\\\\;\\\\/\\$]+)?",
                    RegexOptions.Multiline);

                foreach (var lineResult in matchCollections)
                {
                    var lineResultParse = lineResult.ToString();

                    if (lineResultParse.Equals(string.Empty)) continue;
                    if (lineResultParse.Contains("From -")) i = 0;

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
                        string elementName;
                        elementName = result[0].Replace(" ", "");
                        elementName = elementName.Replace(elementName.Contains("/") ? "/" : " ", "-");
                        string elementValue = result[1];
                        elementValue = elementValue == " " ? "n/a" : result[1];
                        XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                            string.Empty);
                        element.InnerText = elementValue;
                        itemElement.AppendChild(element);
                    }

                    i++;
                }
            }

            xmlDocument.Save($"D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\test\\US Film Order.xml");
            i = 0;
        }

        public static void NewOrder()
        {
            string[] lines = File.ReadAllLines(@"D:\SourceCode\Simple-Text-Parser\Parser\Data\test\New Order4");

            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xmlDocument.DocumentElement;
            xmlDocument.InsertBefore(xmlDeclaration, root);
            XmlElement bodyElement = xmlDocument.CreateElement(string.Empty, "items", string.Empty);
            xmlDocument.AppendChild(bodyElement);

            var i = 0;

            XmlElement itemElement = null;
            try
            {
                foreach (var line in lines)
                {
                    //MatchCollection matchCollections = Regex.Matches(line,
                    //    "(Date|NAME|PORGANIZATION|PADDRESS1|PADDRESS2|PCITY|PSTATE|PZIP|PCOUNT|PPHONE|PFAX|PEMAIL|CCTYPE|CCNUMBER|CCEXPDATE|CCNAME|CCADDRESS1|CCADDRESS2|CCCITY|CCSTATE|CCZIP|CCADDRESS3|COMMENT)(:)?( )?(=)?([ A-za-z0-9\\+\\,\\:\\@\\.\\\"\\-\\\\;\\\\/\\(\\)]+)?",
                    //    RegexOptions.Multiline);

                    MatchCollection matchCollections = Regex.Matches(line,
                        "(Date|email|how|b_firstname|b_lastname|b_address1|b_address2|b_city|b_state|b_zip|b_country|different_shipping|s_firstname|s_lastname|s_address1|s_address2|s_city|s_state|s_zip|s_country|paymethod|card_number|card_holder_name|expiration_dt|button)(:)?( )?(=)?([ A-za-z0-9\\+\\,\\:\\@\\.\\\"\\-\\\\;\\\\/]+)?",
                        RegexOptions.Multiline);
                    foreach (var lineResult in matchCollections)
                    {
                        var lineResultParse = lineResult.ToString();
                        if (lineResultParse.Equals(string.Empty)) continue;

                        if (lineResultParse.Contains("Date"))
                        {
                            i = 0;
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
                                string elementName = result[0].Replace(" ", "");
                                string elementValue = result[1];
                                elementValue = elementValue == string.Empty ? "n/a" : result[1];
                                XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                                    string.Empty);
                                element.InnerText = elementValue;
                                itemElement.AppendChild(element);
                            }
                        }
                        if (lineResultParse.Contains("="))
                        {
                            var result = lineResultParse.Split(new[] { "=" }, StringSplitOptions.None);
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
                                string elementName = result[0].Replace(" ", "");
                                string elementValue = result[1];
                                elementValue = elementValue == " " ? "n/a" : result[1];
                                XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                                    string.Empty);
                                element.InnerText = elementValue;
                                itemElement.AppendChild(element);
                            }
                        }

                        i++;
                    }
                }

                xmlDocument.Save($"D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\test\\New Order4.xml");
                i = 0;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void Orders()
        {
            string[] lines = File.ReadAllLines(@"D:\SourceCode\Simple-Text-Parser\Parser\Data\test\order-bkp");

            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xmlDocument.DocumentElement;
            xmlDocument.InsertBefore(xmlDeclaration, root);
            XmlElement bodyElement = xmlDocument.CreateElement(string.Empty, "items", string.Empty);
            xmlDocument.AppendChild(bodyElement);

            var i = 0;

            XmlElement itemElement = null;

            foreach (var line in lines)
            {
                if (line.Contains("</x-charset>"))
                {
                    i = 0;
                    continue;
                }
                //MatchCollection matchCollections = Regex.Matches(line,
                //    "(Date|NAME|PORGANIZATION|PADDRESS1|PADDRESS2|PCITY|PSTATE|PZIP|PCOUNT|PPHONE|PFAX|PEMAIL|CCTYPE|CCNUMBER|CCEXPDATE|CCNAME|CCADDRESS1|CCADDRESS2|CCCITY|CCSTATE|CCZIP|CCADDRESS3|COMMENT)(:)?( )?(=)?([ A-za-z0-9\\+\\,\\:\\@\\.\\\"\\-\\\\;\\\\/\\(\\)]+)?",
                //    RegexOptions.Multiline);
                string elementName;
                string elementValue;
                MatchCollection matchCollections = Regex.Matches(line,
                    "(Date|Name|Address|Address Cont|City|State|Zip/Postal Code|Country|Credit Card Type|Card Number|Expiration|Card Security Code|Agreement signoff|Phone|Fax|Email|Shipping/Handling)(:)( )([ A-za-z0-9\\+\\,\\:\\@\\.\\\"\\-\\\\;\\\\/\\$\\(\\)]+)?",
                    RegexOptions.Multiline);
                foreach (var lineResult in matchCollections)
                {
                    var lineResultParse = lineResult.ToString();
                    if (lineResultParse.Equals(string.Empty)) continue;
                    if (lineResultParse.Contains("</x-charset>"))
                    {
                        i = 0;
                        continue;
                    }

                    string[] result;
                    if (lineResultParse.Contains("Shipping/Handling"))
                        result = lineResultParse.Split(new[] { "+" }, StringSplitOptions.None);
                    else
                        result = lineResultParse.Split(new[] { ":" }, 2, StringSplitOptions.None); ;
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
                        elementName = result[0].Replace(" ", "-");
                        elementName =
                            elementName.Replace(
                                (elementName.Contains("/") || elementName.Contains("?")) ? "/" : " ", "-");
                        elementValue = result[1];
                        elementValue = elementValue == " " ? "n/a" : result[1];
                        XmlElement element = xmlDocument.CreateElement(string.Empty, elementName.ToLower(),
                            string.Empty);
                        element.InnerText = elementValue;
                        itemElement.AppendChild(element);
                    }

                    i++;
                }
            }

            xmlDocument.Save($"D:\\SourceCode\\Simple-Text-Parser\\Parser\\Data\\test\\order-bkp.xml");
            i = 0;
        }
    }
}