﻿using System;
using System.IO;
using System.Net;
using System.Text;
using System.Xml;

namespace Rubberduck.Common
{
    public enum hAlignment
    {
        Left,
        Center,
        Right
    }

    public enum vAlignment
    {
        Top,
        Middle,
        Bottom
    }

    public class CellFormatting
    {
        public hAlignment HorizontalAlignment;
        public vAlignment VerticalAlignment;
        public string FormatString;
        public bool IsBold;
    }

    public class ColumnInfo
    {
        public ColumnInfo(string title, hAlignment horizontalAlignment = hAlignment.Left, vAlignment verticalAlignment = vAlignment.Bottom)
        {
            Title = title;
            Data = new CellFormatting
            {
                HorizontalAlignment = horizontalAlignment,
                VerticalAlignment = verticalAlignment
            };

            Heading = new CellFormatting
            {
                HorizontalAlignment = horizontalAlignment,
                VerticalAlignment = verticalAlignment
            };
        }
        public CellFormatting Heading;
        public CellFormatting Data;
        public string Title;
    }

    public static class ExportFormatter
    {
        public static string Csv(object[][] data, string title, ColumnInfo[] columnInfos)
        {
            var headerRow = new string[columnInfos.Length];
            for (var c = 0; c < columnInfos.Length; c++)
            {
                headerRow[c] = CsvEncode(columnInfos[c].Title);
            }

            var rows = new string[data.Length];
            for (var r = 0; r < data.Length; r++)
            {
                var row = new string[data[r].Length];
                for (var c = 0; c < data[r].Length; c++)
                {
                    row[c] = CsvEncode(data[r][c]);
                }
                rows[r] = string.Join(",", row);
            }
            return CsvEncode(title.Replace("\r\n"," ")) + Environment.NewLine + string.Join(",", headerRow) + Environment.NewLine + string.Join(Environment.NewLine, rows);
        }

        private static string CsvEncode(object value)
        {
            var s = string.Empty;
            if (value is string)
            {
                s = value.ToString();

                //Escape commas
                if (s.IndexOf(",") >= 0 || s.IndexOf("\"") >= 0)
                {
                    //replace CrLf with Lf
                    s = s.Replace("\r\n", "\n");

                    //escape double-quotes
                    s = "\"" + s.Replace("\"", "\"\"") + "\"";
                }
            }
            else
            {
                if (value != null)
                { 
                    s = value.ToString();
                }
            }
            return s;
        }

        public static string HtmlClipboardFragment(object[][] data, string title, ColumnInfo[] columnInfos)
        {
            const string OffsetFormat = "0000000000";
            const string CFHeaderTemplate = 
                "Version:1.0\r\n" +
                "StartHTML:{0}\r\n" +
                "EndHTML:{1}\r\n" +
                "StartFragment:{2}\r\n" +
                "EndFragment:{3}\r\n";
            
            const string HtmlHeader = 
                "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\r\n" + 
                "<html xmlns=\"http://www.w3.org/1999/xhtml\">\r\n" +
                "<body>\r\n" +
                "<!--StartFragment-->\r\n";

            const string HtmlFooter = 
                "<!--EndFragment-->\r\n" +
                "</body>\r\n" +
                "</html>";

            var html = HtmlTable(data, title, columnInfos);

            var CFHeaderLength = string.Format(CFHeaderTemplate, OffsetFormat, OffsetFormat, OffsetFormat, OffsetFormat).Length;
            var startFragment = CFHeaderLength + HtmlHeader.Length;
            var endFragment = startFragment + html.Length;
            var endHTML = endFragment + HtmlFooter.Length;

            var CfHtml = string.Format(CFHeaderTemplate, CFHeaderLength.ToString(OffsetFormat), endHTML.ToString(OffsetFormat), startFragment.ToString(OffsetFormat), endFragment.ToString(OffsetFormat));

            return CfHtml + HtmlHeader + html + HtmlFooter;
        }

        public static string HtmlTable(object[][] data, string title, ColumnInfo[] columnInfos)
        {            

            var titleRow = HtmlCell(title,true,false,3,columnInfos.Length);

            var hcells = new string[columnInfos.Length];
            for (var c = 0; c < columnInfos.Length; c++)
            {
                hcells[c] = HtmlCell(columnInfos[c].Title, true, true, 3, 1, columnInfos[c].Heading.HorizontalAlignment);
            }
            var headerRow = "  <tr>\r\n" + string.Join(Environment.NewLine, hcells) + "\r\n</tr>";

            var rows = new string[data.Length];
            for (var r = 0; r < data.Length; r++)
            {
                var row = new string[data[r].Length];
                for (var c = 0; c < data[r].Length; c++)
                {
                    row[c] = HtmlCell(data[r][c], r == data.Length - 1, false, 3, 1, columnInfos[c].Heading.HorizontalAlignment);
                }
                rows[r] = "  <tr>\r\n" + string.Join(Environment.NewLine, row) + "\r\n</tr>";
            }
            return  "<table cellspacing=\"0\">\r\n" + titleRow + "\r\n" + headerRow + "\r\n" + string.Join(Environment.NewLine, rows) + "\r\n</table>\r\n";
        }

        private static string HtmlCell(object value, bool bottomBorder = false, bool bold = false, int padding = 3, int colSpan = 1, hAlignment hAlign = hAlignment.Left)
        {
            const string td = "    <td style=\"{0}\"{1}><div style=\"{2}\">{3}</div></td>";
            const string nbsp = "&#160;";

            var cellContent = nbsp;
            var colspanAttribute = colSpan == 1 ? "" : " colspan=\"" + colSpan + "\"";
            var border = bottomBorder ? "0.5pt" : "";
            if (value != null)
            {
                cellContent = value.ToString().HtmlEncode();
            }
            return string.Format(td, TdStyle(hAlign, border, bold), colspanAttribute, TdDivStyle(padding, hAlign), cellContent);
        }

        private static string TdStyle(hAlignment hAlign = hAlignment.Left, string borderBottom = "", bool isBold = false)
        {
            const string tdstyle = "vertical-align: bottom; ";

            var sAlign = $"text-align: {hAlign.ToString()}; " ;
            var sBorder = borderBottom.Length > 0 ? "border-bottom: " + borderBottom + " solid #000000; " : "";
            var sWeight = isBold ? "font-weight: bold; " : "";
            return tdstyle + sAlign + sBorder + sWeight;
        }

        private static string TdDivStyle(int padding, hAlignment hAlign = hAlignment.Left)
        {
            switch (hAlign)
            {
                case hAlignment.Left:
                    return "vertical-align: bottom; padding-left: " + padding + "px; ";
                case hAlignment.Right:
                    return "vertical-align: bottom; padding-right: " + padding + "px; ";
                default:
                    return "vertical-align: bottom; padding-left: " + padding + "px; padding-right: " + padding + "px; ";
            }
        }

        private static string HtmlEncode(this string value)
        {
            return WebUtility.HtmlEncode(value.ToString());
        }

        public static MemoryStream XmlSpreadsheetNew(object[][] data, string title, ColumnInfo[] columnInfos)
        {
            var strm = new MemoryStream();

            var settings = new XmlWriterSettings
            {
                Indent = true,
                Encoding = new UTF8Encoding(false)
            };

            using (var xmlSS = XmlWriter.Create(strm, settings))
            {
                xmlSS.WriteStartDocument();

                //Processing Instructions
                xmlSS.WriteProcessingInstruction("mso-application", "progid=\"Excel.Sheet\"");
                //Namespaces
                xmlSS.WriteStartElement("Workbook", "urn:schemas-microsoft-com:office:spreadsheet");
                xmlSS.WriteAttributeString("xmlns", null, null, "urn:schemas-microsoft-com:office:spreadsheet");
                xmlSS.WriteAttributeString("xmlns", "o", null, "urn:schemas-microsoft-com:office:office");
                xmlSS.WriteAttributeString("xmlns", "x", null, "urn:schemas-microsoft-com:office:excel");
                xmlSS.WriteAttributeString("xmlns", "ss", null, "urn:schemas-microsoft-com:office:spreadsheet");
                xmlSS.WriteAttributeString("xmlns", "html", null, "http://www.w3.org/TR/REC-html40");

                xmlSS.WriteStartElement("Styles");

                //Default Style
                xmlSS.WriteStartElement("Style");
                xmlSS.WriteAttributeString("ss", "ID", null, "Default");
                xmlSS.WriteAttributeString("ss", "Name", null, "Normal");
                xmlSS.WriteStartElement("Alignment");
                xmlSS.WriteAttributeString("ss", "Vertical", null, "Bottom");
                xmlSS.WriteEndElement(); //Close Alignment
                xmlSS.WriteStartElement("Font");
                xmlSS.WriteAttributeString("ss", "FontName", null, "Calibri");
                xmlSS.WriteAttributeString("x", "Family", null, "Swiss");
                xmlSS.WriteAttributeString("ss", "Size", null, "11");
                xmlSS.WriteAttributeString("ss", "Color", null, "#000000");
                xmlSS.WriteEndElement(); //Close Font
                xmlSS.WriteElementString("Interior", "");
                xmlSS.WriteElementString("NumberFormat", "");
                xmlSS.WriteElementString("Protection", "");
                xmlSS.WriteEndElement(); //Close Style

                //Style for column headers
                xmlSS.WriteStartElement("Style");
                xmlSS.WriteAttributeString("ss", "ID", null, "HeaderBottomLeft");

                xmlSS.WriteStartElement("Alignment");
                xmlSS.WriteAttributeString("ss", "Horizontal", null, "Left");
                xmlSS.WriteAttributeString("ss", "Vertical", null, "Bottom");
                xmlSS.WriteEndElement(); //Close Alignment

                xmlSS.WriteStartElement("Borders");

                xmlSS.WriteStartElement("Border");
                xmlSS.WriteAttributeString("ss", "Position", null, "Top");
                xmlSS.WriteAttributeString("ss", "LineStyle", null, "Continuous");
                xmlSS.WriteAttributeString("ss", "Weight", null, "1");
                xmlSS.WriteEndElement(); //Close Border

                xmlSS.WriteStartElement("Border");
                xmlSS.WriteAttributeString("ss", "Position", null, "Bottom");
                xmlSS.WriteAttributeString("ss", "LineStyle", null, "Continuous");
                xmlSS.WriteAttributeString("ss", "Weight", null, "1");
                xmlSS.WriteEndElement(); //Close Border

                xmlSS.WriteEndElement(); //Close Borders

                xmlSS.WriteStartElement("Font");
                xmlSS.WriteAttributeString("ss", "Bold", null, "1");
                xmlSS.WriteEndElement(); //Close Font
                xmlSS.WriteEndElement(); //Close Style

                //Header_BottomRight
                xmlSS.WriteStartElement("Style");
                xmlSS.WriteAttributeString("ss", "ID", null, "HeaderBottomRight");

                xmlSS.WriteStartElement("Alignment");
                xmlSS.WriteAttributeString("ss", "Horizontal", null, "Right");
                xmlSS.WriteAttributeString("ss", "Vertical", null, "Bottom");
                xmlSS.WriteEndElement(); //Close Alignment

                xmlSS.WriteStartElement("Borders");

                xmlSS.WriteStartElement("Border");
                xmlSS.WriteAttributeString("ss", "Position", null, "Top");
                xmlSS.WriteAttributeString("ss", "LineStyle", null, "Continuous");
                xmlSS.WriteAttributeString("ss", "Weight", null, "1");
                xmlSS.WriteEndElement(); //Close Border

                xmlSS.WriteStartElement("Border");
                xmlSS.WriteAttributeString("ss", "Position", null, "Bottom");
                xmlSS.WriteAttributeString("ss", "LineStyle", null, "Continuous");
                xmlSS.WriteAttributeString("ss", "Weight", null, "1");
                xmlSS.WriteEndElement(); //Close Border

                xmlSS.WriteEndElement(); //Close Borders

                xmlSS.WriteStartElement("Font");
                xmlSS.WriteAttributeString("ss", "Bold", null, "1");
                xmlSS.WriteEndElement(); //Close Font
                xmlSS.WriteEndElement(); //Close Style

                //Style for last row
                xmlSS.WriteStartElement("Style");
                xmlSS.WriteAttributeString("ss", "ID", null, "LastRow");
                xmlSS.WriteStartElement("Borders");
                xmlSS.WriteStartElement("Border");
                xmlSS.WriteAttributeString("ss", "Position", null, "Bottom");
                xmlSS.WriteAttributeString("ss", "LineStyle", null, "Continuous");
                xmlSS.WriteAttributeString("ss", "Weight", null, "1");
                xmlSS.WriteEndElement(); //Close Border
                xmlSS.WriteEndElement(); //Close Borders
                xmlSS.WriteEndElement(); //Close Style


                //Style for right-aligned data cells
                xmlSS.WriteStartElement("Style");
                xmlSS.WriteAttributeString("ss", "ID", null, "RightAligned");
                xmlSS.WriteStartElement("Alignment");
                xmlSS.WriteAttributeString("ss", "Horizontal", null, "Right");
                xmlSS.WriteEndElement(); //Close Alignment
                xmlSS.WriteEndElement(); //Close Style

                //Style for right-aligned last row data cells
                xmlSS.WriteStartElement("Style");
                xmlSS.WriteAttributeString("ss", "ID", null, "LastRowRightAligned");
                xmlSS.WriteStartElement("Alignment");
                xmlSS.WriteAttributeString("ss", "Horizontal", null, "Right");
                xmlSS.WriteEndElement(); //Close Alignment
                xmlSS.WriteStartElement("Borders");
                xmlSS.WriteStartElement("Border");
                xmlSS.WriteAttributeString("ss", "Position", null, "Bottom");
                xmlSS.WriteAttributeString("ss", "LineStyle", null, "Continuous");
                xmlSS.WriteAttributeString("ss", "Weight", null, "1");
                xmlSS.WriteEndElement(); //Close Border
                xmlSS.WriteEndElement(); //Close Borders
                xmlSS.WriteEndElement(); //Close Style


                xmlSS.WriteEndElement(); //Close Styles

                xmlSS.WriteStartElement("Worksheet");
                xmlSS.WriteAttributeString("ss", "Name", null, "Sheet1");
                xmlSS.WriteStartElement("Table");
                xmlSS.WriteAttributeString("ss", "ExpandedColumnCount", null, columnInfos.Length.ToString());
                xmlSS.WriteAttributeString("ss", "ExpandedRowCount", null, (data.Length + 2).ToString());
                xmlSS.WriteAttributeString("ss", "DefaultRowHeight", null, "15");

                xmlSS.WriteStartElement("Row");
                xmlSS.WriteStartElement("Cell");
                xmlSS.WriteAttributeString("ss", "MergeAcross", null, (columnInfos.Length - 1).ToString());
                xmlSS.WriteStartElement("Data");
                xmlSS.WriteAttributeString("ss", "Type", null, "String");
                xmlSS.WriteValue(title);
                xmlSS.WriteEndElement(); //Close Data
                xmlSS.WriteEndElement(); //Close Cell

                xmlSS.WriteEndElement(); //Close Row

                //Column Headers
                if (columnInfos.Length > 0)
                {
                    xmlSS.WriteStartElement("Row");
                    foreach (var ch in columnInfos)
                    {
                        xmlSS.WriteStartElement("Cell");
                        xmlSS.WriteAttributeString("ss", "StyleID", null,
                            "Header" + ch.Heading.VerticalAlignment.ToString() +
                            ch.Heading.HorizontalAlignment.ToString());
                        xmlSS.WriteStartElement("Data");
                        xmlSS.WriteAttributeString("ss", "Type", null, "String");
                        xmlSS.WriteValue(ch.Title);
                        xmlSS.WriteEndElement(); //Close Data
                        xmlSS.WriteEndElement(); //Close Cell
                    }

                    xmlSS.WriteEndElement(); //Close Row
                }

                for (var r = 0; r < data.Length; r++)
                {
                    xmlSS.WriteStartElement("Row");
                    for (var c = 0; c < data[r].Length; c++)
                    {
                        var valueType = (data[r][c] is string || data[r][c] == null) ? "String" : "Number";
                        xmlSS.WriteStartElement("Cell");
                        if (columnInfos[c].Data.HorizontalAlignment == hAlignment.Right)
                        {
                            xmlSS.WriteAttributeString("ss", "StyleID", null,
                                (r == data.Length - 1 ? "LastRowRightAligned" : "RightAligned"));
                        }
                        else
                        {
                            if (r == data.Length - 1)
                            {
                                xmlSS.WriteAttributeString("ss", "StyleID", null, "LastRow");
                            }
                        }

                        xmlSS.WriteStartElement("Data");

                        xmlSS.WriteAttributeString("ss", "Type", null, valueType);
                        if (data[r][c] != null)
                        {
                            xmlSS.WriteValue(data[r][c].ToString());
                        }

                        xmlSS.WriteEndElement(); //Close Data
                        xmlSS.WriteEndElement(); //Close Cell
                    }

                    xmlSS.WriteEndElement(); //Close Row
                }

                xmlSS.WriteEndElement(); //Close Table
                xmlSS.WriteEndElement(); //Close Worksheet
                xmlSS.WriteEndElement(); //Close Workbook
                xmlSS.WriteEndDocument();
                xmlSS.Close();

                return strm;
            }
        }
        
        public static string RTF(object[][] data, string title)
        {
            const byte fontSize = 16;    //half-points
            const long colWidth = 1440;  //twips
            const long borderWidth = 10; //twips
            const long cellPadding = 20; //trips

            const string boldFormat = @"\b{{{0}}}\b0";
            const string borderBottom = @"\clbrdrb\brdrw{0}\brdrs";
            const string borderTop = @"\clbrdrt\brdrw{0}\brdrs";
            const string headerFormat = @"\clvertalb{0}\cellx{1}{2}";
            const string cellFormat =   @"\clvertalt{0}\cellx{1}{2}";
            const string cellContent = @"\pard\intbl\ql\sb{0}\sa{0}\li{0}\lr{0}{{{1}}}\cell{2}";
            const string rowStart = @"\trowd\intbl{0}";
            const string rowEnd = @"\row{0}";
            const string rtfStart = @"{{\rtf1{1}\fs{0}{1}";
            const string titleFormat = @"\pard{{{0}}}\par{1}";
            const string rtfEnd = @"}}{0}";

            var newLine = Environment.NewLine;

            var sb = new StringBuilder();
            sb.AppendFormat(rtfStart, fontSize, newLine);
            sb.AppendFormat(titleFormat, title, newLine);

            var cellBorders = string.Format(borderBottom, borderWidth) + string.Format(borderTop, borderWidth);
            for (var r = 0; r < data.Length; r++)
            {
                if (r == 0)
                {
                    sb.AppendFormat(rowStart,newLine);
                    for (var c = 0; c < data[r].Length; c++)
                    {
                        sb.AppendFormat(headerFormat, cellBorders, colWidth * (c+1), newLine);
                    }
                    for (var c = 0; c < data[r].Length; c++)
                    {
                        sb.AppendFormat(cellContent, cellPadding, string.Format(boldFormat,"Col. " + (c+1)), newLine);
                    }
                    sb.AppendFormat(rowEnd,newLine);
                }

                cellBorders = (r == data.Length - 1) ? string.Format(borderBottom, borderWidth) : "";

                sb.AppendFormat(rowStart, newLine);
                for (var c = 0; c < data[r].Length; c++)
                {
                    sb.AppendFormat(cellFormat, cellBorders, colWidth * (c + 1), newLine);
                }
                for (var c = 0; c < data[r].Length; c++)
                {
                    sb.AppendFormat(cellContent, cellPadding, data[r][c], newLine);
                }
                sb.AppendFormat(rowEnd, newLine);
            }
            sb.AppendFormat(rtfEnd, newLine);
            return sb.ToString();
        }
    }
}
