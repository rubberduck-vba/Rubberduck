using Rubberduck.Com.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using static Vanara.PInvoke.OleAut32;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using PARAMDESC = System.Runtime.InteropServices.ComTypes.PARAMDESC;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Com
{
    public readonly struct FuncDescData
    {
        public int Index { get; }
        public string Name { get; }
        public string DocString { get; }
        public int HelpContext { get; }
        public string HelpFile { get; }
        public FUNCDESC FuncDesc { get; }
        public IEnumerable<ParamDescData> Parameters { get; }

        public FuncDescData(int index, string name, string docString, int helpContext, string helpFile, FUNCDESC funcDesc, IEnumerable<ParamDescData> parameters)
        {
            Index = index;
            Name = name;
            DocString = docString;
            HelpContext = helpContext;
            HelpFile = helpFile;
            FuncDesc = funcDesc;
            Parameters = parameters;
        }
    }
}
