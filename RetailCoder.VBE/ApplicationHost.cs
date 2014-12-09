using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;


namespace Rubberduck
{

    public enum HostApplicationType { Unknown, Excel, Access, Word }

    [ComVisible(false)]
    public static class HostApplication
    {
        public static HostApplicationType Type { get; set; }

        public static string Name
        {
            get
            {
                switch (Type)
                {
                    case HostApplicationType.Access:
                        return "Access";
                    case HostApplicationType.Excel:
                        return "Excel";
                    case HostApplicationType.Word:
                        return "Word";
                    default:
                        return "Unknown";
                }
            }
        }
    }
}
