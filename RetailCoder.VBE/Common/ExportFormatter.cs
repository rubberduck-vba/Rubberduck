using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Common
{
    public static class ExportFormatter
    {
        public static string Csv(object[][] data, string Title)
        {
            string s;
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
            const string offsetFormat = "0000000000";
            const string CfHtmlHeader = 
                "Version:1.0\r\n" +
                "StartHTML:{0}\r\n" +
                "EndHTML:{1}\r\n" +
                "StartFragment:{2}\r\n" +
                "EndFragment:{3}\r\n";
            
            const string HtmlHeader = 
                "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\r\n" + 
                "<html xmlns=\"http://www.w3.org/1999/xhtml\">\r\n" +
                "<body>\r\n" +
                "<!--StartFragment-->";

            const string HtmlFooter = 
                "<!--EndFragment-->\r\n" +
                "</body>\r\n" +
                "</html>";

            string html = Html(data,Title);

            int startHTML = string.Format(CfHtmlHeader, offsetFormat, offsetFormat, offsetFormat, offsetFormat).Length;
            int startFragment = HtmlHeader.Length;
            int endFragment = startFragment + html.Length;
            int endHTML = endFragment + HtmlFooter.Length;

            string CfHtml = string.Format(CfHtmlHeader, startHTML.ToString(offsetFormat), startHTML.ToString(offsetFormat), startHTML.ToString(offsetFormat), startHTML.ToString(offsetFormat));

            return CfHtml + HtmlHeader + html + HtmlFooter;
        }

        public static string Html(object[][] data, string Title)
        {            
            string[] rows = new string[data.Length];
            for (var r = 0; r < data.Length; r++)
            {
                string[] row = new string[data[r].Length];
                for (var c = 0; c < data[r].Length; c++)
                {
                    row[c] = HtmlCell(data[r][c], r == data.Length - 1, c == 0 ? 5: 10);
                }
                rows[r] = "  <tr>" + string.Join(Environment.NewLine, row) + "</tr>";
            }
            return  "<table cellspacing=\"0\">" + string.Join(Environment.NewLine, rows) + "</table>";
        }

        private static string HtmlCell(object value, bool BottomBorder = false, int LeftMargin = 10)
        {
            //TODO do XHTML encoding
            const string td = "    <td style=\"{0}\"><div style=\"{1}\">{2}</div></td>";
            const string nbsp = "&#160;";
            string s = "";
            bool AlignLeft = true;
            string Border = BottomBorder ? "0.5pt" : "";
            if (value == null)
            {
                s = nbsp;
            }
            else
            {
                s = value.ToString();
                AlignLeft = value is string;
            }
            return string.Format(td, TdStyle(AlignLeft, Border), TdDivStyle(LeftMargin));
        }
        private static string TdStyle(bool AlignLeft = true, string BorderBottom = "")
        {
            const string tdstyle = "vertical-align:bottom;";

            string sAlign = AlignLeft ? "text-align:left;" : "text-align:right";
            string sBorder = BorderBottom.Length > 0 ? "border-bottom: 1px solid #000000;" : "";

            return tdstyle + sAlign + sBorder;
        }

        private static string TdDivStyle(int LeftMargin)
        {
            return "margin-left:" + LeftMargin + "px;";
        }
    }
}
