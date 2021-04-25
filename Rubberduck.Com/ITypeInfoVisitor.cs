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
    public interface ITypeInfoVisitor
    {
        void VisitTypeDocumentation(string typeName, string docString, int helpContext, string helpFile);
        void VisitTypeDocumentation2(string helpString, int helpStringcontext, string helpStringdll);
        void VisitTypeAttr(TYPEATTR attr);
        void VisitTypeCustData(Guid guid, object value);
        VisitDirectives VisitTypeImplementedTypeInfo(int href, ITypeInfo implementedTypeInfo);
        void VisitTypeImplTypeCustData(Guid guid, object value);
        void VisitTypeVarDesc(VarDescData varDescData);
        void VisitTypeVarCustData(Guid guid, object value);
        void VisitTypeFuncDesc(FuncDescData funcDescData);
        VisitDirectives VisitTypeFuncParameter(ParamDescData paramDescData);
        void VisitTypeFuncCustData(Guid guid, object value);
        IEnumerable<ITypeLibVisitor> ProvideTypeLibVisitors();
    }
}
