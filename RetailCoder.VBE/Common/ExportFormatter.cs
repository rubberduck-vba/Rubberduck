using System;
using System.Net;

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
    }
}
