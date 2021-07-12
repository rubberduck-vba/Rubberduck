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
    public readonly struct ParamDescData
    {
        public int Index { get; }
        public string Name { get; }
        public ELEMDESC ElemDesc { get; }
        public bool IsArray { get; }
        public bool IsReferencedType { get; }
        public bool IsParamArray { get; }
        public bool IsOptional { get; }
        public VarEnum VarEnum { get; }
        public object DefaultValue { get; }
        public ITypeInfo ReferencedTypeInfo { get; }
        public COMException ReferencedTypeInfoException { get; }
        public IDictionary<Guid, object> CustomData { get; }

        public ParamDescData(int index, string name, ELEMDESC elemDesc, bool isArray, bool isReferencedType, bool isParamArray, bool isOptional, VarEnum varEnum, object defaultValue, ITypeInfo referencedTypeInfo, COMException referencedTypeInfoException, IDictionary<Guid, object> customData)
        {
            Index = index;
            Name = name;
            ElemDesc = elemDesc;
            IsArray = isArray;
            IsReferencedType = isReferencedType;
            IsParamArray = isParamArray;
            IsOptional = isOptional;
            VarEnum = varEnum;
            DefaultValue = defaultValue;
            ReferencedTypeInfo = referencedTypeInfo;
            ReferencedTypeInfoException = referencedTypeInfoException;
            CustomData = customData;
        }
    }
}
