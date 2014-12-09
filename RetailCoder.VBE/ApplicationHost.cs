using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;


namespace Rubberduck
{

    public enum HostApp {Unknown, Excel, Access, Word}

    [ComVisible(false)]
    public static class ApplicationHost
    {
        public static HostApp Type { get; set; }

        public static string Name()
        {
            switch (Type)
            {
                case HostApp.Access:
                    return "Access";
                case HostApp.Excel:
                    return "Excel";
                case HostApp.Word:
                    return "Word";
                default:
                    return "Unknown";
            }
        }
    }
}
