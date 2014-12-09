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
        public static HostApp HostApplicationType { get; set; }
    }
}
