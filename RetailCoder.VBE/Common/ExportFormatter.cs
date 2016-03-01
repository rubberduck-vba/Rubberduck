using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Common
{
    public static class ExportFormatter
    {
        public static string Csv(object[][] data, string Title)
        {
            string s = "";
            string[] rows = new string[data.Length];
            for (var r = 0; r < data.Length; r++)
            {
                string[] row = new string[data[r].Length];
                for (var c = 0; c < data[r].Length; c++)
                {
                    row[c] = CsvEncode(data[r][c]);
                }
                rows[r] = string.Join(",", row);
            }
            return CsvEncode(Title) + Environment.NewLine + string.Join(Environment.NewLine, rows);
        }

        private static string CsvEncode(object value)
        {
            string s = "";
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

        public static string HtmlClipboardFragment(object[][] data, string Title)
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

            string html = ExportFormatter.HtmlTable(data,Title);

            int CFHeaderLength = string.Format(CFHeaderTemplate, OffsetFormat, OffsetFormat, OffsetFormat, OffsetFormat).Length;
            int startFragment = CFHeaderLength + HtmlHeader.Length;
            int endFragment = startFragment + html.Length;
            int endHTML = endFragment + HtmlFooter.Length;

            string CfHtml = string.Format(CFHeaderTemplate, CFHeaderLength.ToString(OffsetFormat), endHTML.ToString(OffsetFormat), startFragment.ToString(OffsetFormat), endFragment.ToString(OffsetFormat));

            return CfHtml + HtmlHeader + html + HtmlFooter;
        }

        public static string HtmlTable(object[][] data, string Title)
        {            
            string[] rows = new string[data.Length];
            for (var r = 0; r < data.Length; r++)
            {
                string[] row = new string[data[r].Length];
                for (var c = 0; c < data[r].Length; c++)
                {
                    row[c] = HtmlCell(data[r][c], r == data.Length - 1, c == 0 ? 5: 10);
                }
                rows[r] = "  <tr>\r\n" + string.Join(Environment.NewLine, row) + "\r\n</tr>";
            }
            return  "<table cellspacing=\"0\">\r\n" + string.Join(Environment.NewLine, rows) + "\r\n</table>\r\n";
        }

        private static string HtmlCell(object value, bool BottomBorder = false, int LeftPadding = 10)
        {
            const string td = "    <td style=\"{0}\"><div style=\"{1}\">{2}</div></td>";
            const string nbsp = "&#160;";

            string CellContent = nbsp;
            bool AlignLeft = true;
            string Border = BottomBorder ? "0.5pt" : "";
            if (value != null)
            {
                CellContent = value.ToString().HtmlEncode();
                AlignLeft = value is string;
            }
            return string.Format(td, TdStyle(AlignLeft, Border), TdDivStyle(LeftPadding), CellContent);
        }

        private static string TdStyle(bool AlignLeft = true, string BorderBottom = "")
        {
            const string tdstyle = "vertical-align: bottom; ";

            string sAlign = AlignLeft ? "text-align: left; " : "text-align: right; ";
            string sBorder = BorderBottom.Length > 0 ? "border-bottom: " + BorderBottom + " solid #000000; " : "";

            return tdstyle + sAlign + sBorder;
        }

        private static string TdDivStyle(int LeftPadding)
        {
            return "vertical-align: bottom; padding-left: " + LeftPadding + "px; ";
        }

        private static string HtmlEncode(this string value)
        {
            return WebUtility.HtmlEncode(value.ToString());
        }

        public static string XmlSpreadsheet(object[][] data, string Title)
        {
            StringBuilder s = new StringBuilder();
            s.AppendLine("<?xml version=\"1.0\"?>");
            s.AppendLine("<?mso-application progid=\"Excel.Sheet\"?>");
            s.AppendLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            s.AppendLine(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            s.AppendLine(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"");
            s.AppendLine(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            s.AppendLine(" xmlns:html=\"http://www.w3.org/TR/REC-html40\">");

            s.AppendLine(" <Styles>");
            s.AppendLine("  <Style ss:ID=\"Default\" ss:Name=\"Normal\">");
            s.AppendLine("   <Alignment ss:Vertical=\"Bottom\"/>");
            s.AppendLine("   <Borders/>");
            s.AppendLine("   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/>");
            s.AppendLine("   <Interior/>");
            s.AppendLine("   <NumberFormat/>");
            s.AppendLine("   <Protection/>");
            s.AppendLine("  </Style>");

            s.AppendLine("  <Style ss:ID=\"HeadingLeft\">");
            s.AppendLine("   <Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Bottom\"/>");
            s.AppendLine("   <Borders>");
            s.AppendLine("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            s.AppendLine("    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            s.AppendLine("   </Borders>");
            s.AppendLine("   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/>");
            s.AppendLine("  </Style>");

            s.AppendLine("  <Style ss:ID=\"HeadingRight\">");
            s.AppendLine("   <Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>");
            s.AppendLine("   <Borders>");
            s.AppendLine("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            s.AppendLine("    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            s.AppendLine("   </Borders>");
            s.AppendLine("   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/>");
            s.AppendLine("  </Style>");

            s.AppendLine("  <Style ss:ID=\"LastLeft\">");
            s.AppendLine("   <Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Bottom\"/>");
            s.AppendLine("   <Borders>");
            s.AppendLine("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            s.AppendLine("   </Borders>");
            s.AppendLine("  </Style>");

            s.AppendLine("  <Style ss:ID=\"LastRight\">");
            s.AppendLine("   <Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>");
            s.AppendLine("   <Borders>");
            s.AppendLine("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
            s.AppendLine("   </Borders>");
            s.AppendLine("  </Style>");

            s.AppendLine(" </Styles>");

            s.AppendLine(" <Worksheet ss:Name=\"Sheet1\">");

//            s.AppendLine("  <Table ss:ExpandedColumnCount=\"2\" ss:ExpandedRowCount=\"1\" ss:DefaultRowHeight=\"15\">");

            if (data.Length == 0)
            {
                s.AppendFormat("  <Table ss:ExpandedColumnCount=\"{0}\" ss:ExpandedRowCount=\"{1}\" ss:DefaultRowHeight=\"15\">\r\n", 1, 1);
                //Title 
                s.AppendLine("    <Row>");
                s.AppendFormat("     <Cell ss:StyleID=\"{0}\"><Data ss:Type=\"{1}\">{2}</Data></Cell>\r\n", "Default", "String", Title.HtmlEncode());
                s.AppendLine("    </Row>");
            }
            else
            {
                //Title 
                s.AppendFormat("  <Table ss:ExpandedColumnCount=\"{0}\" ss:ExpandedRowCount=\"{1}\" ss:DefaultRowHeight=\"15\">\r\n", data[0].Length, data.Length + 2);
                s.AppendLine("    <Row>");
                s.AppendFormat("     <Cell ss:MergeAcross=\"{0}\" ss:StyleID=\"{1}\"><Data ss:Type=\"{2}\">{3}</Data></Cell>\r\n", data[0].Length - 1, "Default", "String", Title.HtmlEncode());
                s.AppendLine("    </Row>");

                //Column Headers
                s.AppendLine("    <Row>");
                for (var c = 0; c < data[0].Length; c++)
                {
                    s.AppendFormat("     <Cell ss:StyleID=\"{0}\"><Data ss:Type=\"{1}\">{2}</Data></Cell>\r\n", "HeadingLeft", "String", "Col. " + (c+1).ToString());
                }
                s.AppendLine("    </Row>");
            
                //Data Rows
                string sValue = "";
                string sStyle = "";
                for (var r = 0; r < data.Length; r++)
                {
                    s.AppendLine("    <Row>");
                    for (var c = 0; c < data[r].Length; c++)
                    {
                        sValue = data[r][c] != null ? data[r][c].ToString().HtmlEncode() : "BLANK";
                        
                        sStyle = r == (data.Length - 1) ? "LastLeft" : "Default";
                        s.AppendFormat("     <Cell ss:StyleID=\"{0}\"><Data ss:Type=\"{1}\">{2}</Data></Cell>\r\n", sStyle, "String", sValue);
                    }
                    s.AppendLine("    </Row>");
                }
            }


//            s.AppendLine("    <Row>");
//            s.AppendLine("     <Cell><Data ss:Type=\"String\">Smith</Data></Cell>");
//            s.AppendLine("     <Cell><Data ss:Type=\"String\">Jones</Data></Cell>");
//            s.AppendLine("    </Row>");
            
            s.AppendLine("  </Table>");
            s.AppendLine(" </Worksheet>");
            s.AppendLine("</Workbook>");

            return s.ToString();
        }


        public static string RTF(object[][] data, string Title)
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

            string newLine = Environment.NewLine;

            StringBuilder s = new StringBuilder();
            s.AppendFormat(rtfStart, fontSize.ToString(), newLine);
            s.AppendFormat(titleFormat, Title, newLine);

            string cellBorders = string.Format(borderBottom, borderWidth) + string.Format(borderTop, borderWidth);
            for (var r = 0; r < data.Length; r++)
            {
                if (r == 0)
                {
                    s.AppendFormat(rowStart,newLine);
                    for (int c = 0; c < data[r].Length; c++)
                    {
                        s.AppendFormat(headerFormat, cellBorders, colWidth * (c+1), newLine);
                    }
                    for (int c = 0; c < data[r].Length; c++)
                    {
                        s.AppendFormat(cellContent, cellPadding, string.Format(boldFormat,"Col. " + (c+1).ToString()), newLine);
                    }
                    s.AppendFormat(rowEnd,newLine);
                }

                cellBorders = (r == data.Length - 1) ? string.Format(borderBottom, borderWidth) : "";

                s.AppendFormat(rowStart, newLine);
                for (int c = 0; c < data[r].Length; c++)
                {
                    s.AppendFormat(cellFormat, cellBorders, colWidth * (c + 1), newLine);
                }
                for (int c = 0; c < data[r].Length; c++)
                {
                    s.AppendFormat(cellContent, cellPadding, data[r][c], newLine);
                }
                s.AppendFormat(rowEnd, newLine);
            }
            s.AppendFormat(rtfEnd, newLine);
            return s.ToString();
        }
    }
}
