using System;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Utility
{
    internal static class TypeKindMapper
    {
        public static TYPEKIND_VBE ToTypeKindVbe(this TYPEKIND typeKind)
        {
            switch (typeKind)
            {
                case TYPEKIND.TKIND_ALIAS:
                    return TYPEKIND_VBE.TKIND_ALIAS;
                case TYPEKIND.TKIND_UNION:
                    return TYPEKIND_VBE.TKIND_UNION;
                case TYPEKIND.TKIND_COCLASS:
                    return TYPEKIND_VBE.TKIND_COCLASS;
                case TYPEKIND.TKIND_DISPATCH:
                    return TYPEKIND_VBE.TKIND_DISPATCH;
                case TYPEKIND.TKIND_ENUM:
                    return TYPEKIND_VBE.TKIND_ENUM;
                case TYPEKIND.TKIND_INTERFACE:
                    return TYPEKIND_VBE.TKIND_INTERFACE;
                case TYPEKIND.TKIND_MODULE:
                    return TYPEKIND_VBE.TKIND_MODULE;
                case TYPEKIND.TKIND_RECORD:
                    return TYPEKIND_VBE.TKIND_RECORD;

                case TYPEKIND.TKIND_MAX:
                    return TYPEKIND_VBE.TKIND_VBACLASS;
                default:
                    throw new NotSupportedException(typeKind.ToString());
            }
        }
    }
}