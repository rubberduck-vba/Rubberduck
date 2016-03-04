using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using FUNCFLAGS = System.Runtime.InteropServices.ComTypes.FUNCFLAGS;
using FUNCKIND = System.Runtime.InteropServices.ComTypes.FUNCKIND;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Parsing.Symbols
{
    public class ReferencedDeclarationsCollector
    {
        /// <summary>
        /// Controls how a type library is registered.
        /// </summary>
        private enum REGKIND
        {
            /// <summary>
            /// Use default register behavior.
            /// </summary>
            REGKIND_DEFAULT = 0,
            /// <summary>
            /// Register this type library.
            /// </summary>
            REGKIND_REGISTER = 1,
            /// <summary>
            /// Do not register this type library.
            /// </summary>
            REGKIND_NONE = 2
        }

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
        private static extern void LoadTypeLibEx(string strTypeLibName, REGKIND regKind, out ITypeLib TypeLib);

        private static readonly IDictionary<VarEnum, string> TypeNames = new Dictionary<VarEnum, string>
        {
            {VarEnum.VT_DISPATCH, "DISPATCH"},
            {VarEnum.VT_VOID, string.Empty},
            {VarEnum.VT_VARIANT, "Variant"},
            {VarEnum.VT_BLOB_OBJECT, "Object"},
            {VarEnum.VT_STORED_OBJECT, "Object"},
            {VarEnum.VT_STREAMED_OBJECT, "Object"},
            {VarEnum.VT_BOOL, "Boolean"},
            {VarEnum.VT_BSTR, "String"},
            {VarEnum.VT_LPSTR, "String"},
            {VarEnum.VT_LPWSTR, "String"},
            {VarEnum.VT_I1, "Variant"}, // no signed byte type in VBA
            {VarEnum.VT_UI1, "Byte"},
            {VarEnum.VT_I2, "Integer"},
            {VarEnum.VT_UI2, "Variant"}, // no unsigned integer type in VBA
            {VarEnum.VT_I4, "Long"},
            {VarEnum.VT_UI4, "Variant"}, // no unsigned long integer type in VBA
            {VarEnum.VT_I8, "Variant"}, // LongLong on 64-bit VBA
            {VarEnum.VT_UI8, "Variant"}, // no unsigned LongLong integer type in VBA
            {VarEnum.VT_INT, "Long"}, // same as I4
            {VarEnum.VT_UINT, "Variant"}, // same as UI4
            {VarEnum.VT_DATE, "Date"},
            {VarEnum.VT_DECIMAL, "Currency"}, // best match?
            {VarEnum.VT_EMPTY, "Empty"},
            {VarEnum.VT_R4, "Single"},
            {VarEnum.VT_R8, "Double"},
        };

        public void DebugOutputAllReferences(References references)
        {
            foreach (var reference in references.Cast<Reference>().Where(reference => !reference.IsBroken))
            {
                try
                {
                    var projectName = reference.Name;
                    var path = reference.FullPath;

                    Debug.WriteLine("Project name: {0} ({1})", projectName, path);

                    ITypeLib typeLibrary;
                    LoadTypeLibEx(path, REGKIND.REGKIND_NONE, out typeLibrary);

                    var typeCount = typeLibrary.GetTypeInfoCount();
                    for (var i = 0; i < typeCount; i++)
                    {
                        ITypeInfo info;
                        typeLibrary.GetTypeInfo(i, out info);

                        if (info == null)
                        {
                            continue;
                        }

                        string typeName;
                        string docString;
                        int helpContext;
                        string helpFile;
                        info.GetDocumentation(-1, out typeName, out docString, out helpContext, out helpFile);

                        TYPEKIND typeKind;
                        typeLibrary.GetTypeInfoType(i, out typeKind);

                        DeclarationType typeDeclarationType = DeclarationType.Control; // todo: a better default
                        if (typeKind == TYPEKIND.TKIND_ENUM)
                        {
                            typeDeclarationType = DeclarationType.Enumeration;
                        }
                        else if (typeKind == TYPEKIND.TKIND_COCLASS || typeKind == TYPEKIND.TKIND_INTERFACE || typeKind == TYPEKIND.TKIND_ALIAS || typeKind == TYPEKIND.TKIND_DISPATCH)
                        {
                            typeDeclarationType = DeclarationType.Class;
                        }
                        else if (typeKind == TYPEKIND.TKIND_RECORD)
                        {
                            typeDeclarationType = DeclarationType.UserDefinedType;
                        }
                        else if (typeKind == TYPEKIND.TKIND_MODULE)
                        {
                            typeDeclarationType = DeclarationType.Module;
                        }


                        Debug.WriteLine("Type name: {0} ({1}) {2}", typeName, typeDeclarationType, docString);

                        IntPtr ppTypeAttr;
                        info.GetTypeAttr(out ppTypeAttr);

                        var implements = new List<string>();
                        var typeAttr = (TYPEATTR)Marshal.PtrToStructure(ppTypeAttr, typeof (TYPEATTR));
                        for (var implIndex = 0; implIndex < typeAttr.cImplTypes; implIndex++)
                        {
                            int href;
                            info.GetRefTypeOfImplType(implIndex, out href);

                            ITypeInfo implTypeInfo;
                            info.GetRefTypeInfo(href, out implTypeInfo);

                            string implTypeName;
                            string implDocString;
                            int implHelpContext;
                            string implHelpFile;
                            implTypeInfo.GetDocumentation(-1, out implTypeName, out implDocString, out implHelpContext, out implHelpFile);
                            
                            implements.Add(implTypeName);
                            Debug.WriteLine("\tImplements {0} {1}", implTypeName, implDocString);
                        }

                        for (var funcIndex = 0; funcIndex < typeAttr.cFuncs; funcIndex++)
                        {
                            IntPtr ppFuncDesc;
                            info.GetFuncDesc(funcIndex, out ppFuncDesc);
                            var funcDesc = (FUNCDESC) Marshal.PtrToStructure(ppFuncDesc, typeof (FUNCDESC));
                            
                            var names = new string[255];
                            int namesArrayLength;
                            info.GetNames(funcDesc.memid, names, 255, out namesArrayLength);

                            var memberName = names[0];

                            var funcValueType = (VarEnum)funcDesc.elemdescFunc.tdesc.vt;
                            DeclarationType memberType;
                            if (funcDesc.invkind.HasFlag(INVOKEKIND.INVOKE_PROPERTYGET))
                            {
                                memberType = DeclarationType.PropertyGet;
                            }
                            else if (funcDesc.invkind.HasFlag(INVOKEKIND.INVOKE_PROPERTYPUT))
                            {
                                memberType = DeclarationType.PropertyLet;
                            }
                            else if (funcDesc.invkind.HasFlag(INVOKEKIND.INVOKE_PROPERTYPUTREF))
                            {
                                memberType = DeclarationType.PropertySet;
                            }
                            else if (funcValueType == VarEnum.VT_VOID)
                            {
                                memberType = DeclarationType.Procedure;
                            }
                            else if (funcDesc.funckind == FUNCKIND.FUNC_PUREVIRTUAL)
                            {
                                memberType = DeclarationType.Event;
                            }
                            else
                            {
                                memberType = DeclarationType.Function;
                            }

                            var asTypeName = string.Empty;
                            if (memberType != DeclarationType.Procedure && !TypeNames.TryGetValue(funcValueType, out asTypeName))
                            {
                                asTypeName = funcValueType.ToString(); //TypeNames[VarEnum.VT_VARIANT];
                            }

                            Debug.WriteLine("Member: {0} ({1}) :: {2}", memberName, memberType, asTypeName);
                        }

                        for (var fieldIndex = 0; fieldIndex < typeAttr.cVars; fieldIndex++)
                        {
                            IntPtr ppVarDesc;
                            info.GetVarDesc(fieldIndex, out ppVarDesc);

                            var varDesc = (VARDESC) Marshal.PtrToStructure(ppVarDesc, typeof (VARDESC));

                            var names = new string[255];
                            int namesArrayLength;
                            info.GetNames(varDesc.memid, names, 255, out namesArrayLength);

                            var fieldName = names[0];
                            var fieldValueType = (VarEnum)varDesc.elemdescVar.tdesc.vt;

                            var memberType = DeclarationType.Variable;
                            if (varDesc.varkind == VARKIND.VAR_CONST)
                            {
                                memberType = typeDeclarationType == DeclarationType.Enumeration
                                    ? DeclarationType.EnumerationMember 
                                    : DeclarationType.Constant;
                            }
                            else if (typeDeclarationType == DeclarationType.UserDefinedType)
                            {
                                memberType = DeclarationType.UserDefinedTypeMember;
                            }

                            string asTypeName;
                            if (!TypeNames.TryGetValue(fieldValueType, out asTypeName))
                            {
                                asTypeName = TypeNames[VarEnum.VT_VARIANT];
                            }

                            Debug.WriteLine("Field: {0} ({1}) :: {2}", fieldName, memberType, asTypeName);
                        }
                    }
                    
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }
    }
}
