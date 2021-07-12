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
    public class TypeLibWalker : TypeApiWalker<ITypeLibVisitor>
    {
        private readonly ITypeLib _typeLib;
        private readonly ITypeLib2 _typeLib2;

        private TypeLibWalker(ITypeLib typeLib, IEnumerable<ITypeLibVisitor> visitors)
        {
            _typeLib = typeLib;
            _typeLib2 = typeLib as ITypeLib2;
            Visitors = visitors;
        }

        public static void Accept(ITypeLib typeLib, ITypeLibVisitor visitor)
        {
            Accept(typeLib, new[] { visitor });
        }

        public static void Accept(ITypeLib typeLib, IEnumerable<ITypeLibVisitor> visitors)
        {
            var walker = new TypeLibWalker(typeLib, visitors);
            walker.WalkTypeLib();
        }

        private void WalkTypeLib()
        {
            _typeLib.GetDocumentation(WellKnown.MemberIds.MEMBERID_NIL, out var libName, out var docString, out var helpContext, out var helpFile);
            ExecuteVisit(v => v.VisitLibDocumentation(libName, docString, helpContext, helpFile, _typeLib.GetTypeInfoCount()));

            _typeLib.UsingLibAttr(attr => ExecuteVisit(v => v.VisitLibAttr(attr)));

            if (_typeLib2 != null)
            {
                _typeLib2.GetDocumentation2(WellKnown.MemberIds.MEMBERID_NIL, out var helpString, out var helpStringContext, out var helpStringdll);
                ExecuteVisit(v => v.VisitLibDocumentation2(helpString, helpStringContext, helpStringdll));
                EnumerateCustomData(
                    _typeLib2.GetAllCustData,
                    (guid, value) => ExecuteVisit(v => v.VisitLibCustData(guid, value)));
            }

            for (var i = 0; i < _typeLib.GetTypeInfoCount(); i++)
            {
                _typeLib.GetTypeInfo(i, out var typeInfo);
                TypeInfoWalker.Accept(_typeLib, i, typeInfo, Visitors.SelectMany(v => v.ProvideTypeInfoVisitors()));
            }
        }
    }
}
