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
    internal struct VarDescDataInternal
    {
        public bool IsArray { get; set; }
        public bool IsReferencedType { get; set; }
        public VarEnum VarEnum { get; set; }
        public ITypeInfo ReferencedTypeInfo { get; set; }
    }

    public readonly struct VarDescData
    {
        public string Name { get; }
        public VARDESC VarDesc { get; }
        public bool IsArray { get; }
        public bool IsReferencedType { get; }
        public VarEnum VarEnum { get; }
        public ITypeInfo ReferencedTypeInfo { get; }
        public object Value { get; }

        internal VarDescData(string name, VARDESC varDesc, VarDescDataInternal finalDesc, object value)
        {
            Name = name;
            VarDesc = varDesc;
            IsArray = finalDesc.IsArray;
            IsReferencedType = finalDesc.IsReferencedType;
            VarEnum = finalDesc.VarEnum;
            ReferencedTypeInfo = finalDesc.ReferencedTypeInfo;
            Value = value;
        }
    }
}
