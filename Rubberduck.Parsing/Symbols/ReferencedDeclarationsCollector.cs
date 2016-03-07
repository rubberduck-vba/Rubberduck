using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;
using FUNCFLAGS = System.Runtime.InteropServices.ComTypes.FUNCFLAGS;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using FUNCKIND = System.Runtime.InteropServices.ComTypes.FUNCKIND;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
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

        private string GetTypeName(ITypeInfo info)
        {
            string typeName;
            string docString; // todo: put the docString to good use?
            int helpContext;
            string helpFile;
            info.GetDocumentation(-1, out typeName, out docString, out helpContext, out helpFile);

            return typeName;
        }

        public IEnumerable<Declaration> GetDeclarationsForReference(Reference reference)
        {
            var projectName = reference.Name;
            var path = reference.FullPath;

            var projectQualifiedModuleName = new QualifiedModuleName(projectName, projectName);
            var projectQualifiedMemberName = new QualifiedMemberName(projectQualifiedModuleName, projectName);

            var projectDeclaration = new Declaration(projectQualifiedMemberName, null, null, projectName, false, false, Accessibility.Global, DeclarationType.Project);
            yield return projectDeclaration;

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

                var typeName = GetTypeName(info);
                var typeDeclarationType = GetDeclarationType(typeLibrary, i);

                QualifiedModuleName typeQualifiedModuleName;
                QualifiedMemberName typeQualifiedMemberName;
                if (typeDeclarationType == DeclarationType.Enumeration ||
                    typeDeclarationType == DeclarationType.UserDefinedType)
                {
                    typeQualifiedModuleName = projectQualifiedModuleName;
                    typeQualifiedMemberName = new QualifiedMemberName(projectQualifiedModuleName, typeName);
                }
                else
                {
                    typeQualifiedModuleName = new QualifiedModuleName(projectName, typeName);
                    typeQualifiedMemberName = new QualifiedMemberName(typeQualifiedModuleName, typeName);
                }

                var moduleDeclaration = new Declaration(typeQualifiedMemberName, projectDeclaration, projectDeclaration, typeName, false, false, Accessibility.Global, typeDeclarationType, null, Selection.Home);
                yield return moduleDeclaration;

                IntPtr typeAttributesPointer;
                info.GetTypeAttr(out typeAttributesPointer);

                var typeAttributes = (TYPEATTR)Marshal.PtrToStructure(typeAttributesPointer, typeof (TYPEATTR));
                //var implements = GetImplementedInterfaceNames(typeAttributes, info);

                for (var memberIndex = 0; memberIndex < typeAttributes.cFuncs; memberIndex++)
                {
                    IntPtr memberDescriptorPointer;
                    info.GetFuncDesc(memberIndex, out memberDescriptorPointer);
                    var memberDescriptor = (FUNCDESC) Marshal.PtrToStructure(memberDescriptorPointer, typeof (FUNCDESC));
                            
                    var memberNames = new string[255]; // member name at index 0; array contains parameter names too
                    int namesArrayLength;
                    info.GetNames(memberDescriptor.memid, memberNames, 255, out namesArrayLength);

                    var memberName = memberNames[0];

                    var funcValueType = (VarEnum)memberDescriptor.elemdescFunc.tdesc.vt;
                    var memberDeclarationType = GetDeclarationType(memberDescriptor, funcValueType);

                    var asTypeName = string.Empty;
                    if (memberDeclarationType != DeclarationType.Procedure && !TypeNames.TryGetValue(funcValueType, out asTypeName))
                    {
                        asTypeName = funcValueType.ToString(); //TypeNames[VarEnum.VT_VARIANT];
                    }

                    var memberDeclaration = new Declaration(new QualifiedMemberName(typeQualifiedModuleName, memberName), moduleDeclaration, moduleDeclaration, asTypeName, false, false, Accessibility.Global, memberDeclarationType, null, Selection.Home);
                    yield return memberDeclaration;

                    var parameterCount = memberDescriptor.cParams - 1;
                    for (var paramIndex = 0; paramIndex < parameterCount; paramIndex++)
                    {
                        var paramName = memberNames[paramIndex + 1];

                        var paramPointer = new IntPtr(memberDescriptor.lprgelemdescParam.ToInt64() + Marshal.SizeOf(typeof (ELEMDESC))*paramIndex);
                        var elementDesc = (ELEMDESC) Marshal.PtrToStructure(paramPointer, typeof (ELEMDESC));
                        var isOptional = elementDesc.desc.paramdesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FOPT);
                        var asParamTypeName = string.Empty;
                        
                        var isByRef = false;
                        var isArray = false;
                        var paramDesc = elementDesc.tdesc;
                        var valueType = (VarEnum) paramDesc.vt;
                        if (valueType == VarEnum.VT_PTR || valueType == VarEnum.VT_BYREF)
                        {
                            //var paramTypeDesc = (TYPEDESC) Marshal.PtrToStructure(paramDesc.lpValue, typeof (TYPEDESC));
                            isByRef = true;
                            var paramValueType = (VarEnum) paramDesc.vt;
                            if (!TypeNames.TryGetValue(paramValueType, out asParamTypeName))
                            {
                                asParamTypeName = TypeNames[VarEnum.VT_VARIANT];
                            }
                            //var href = paramDesc.lpValue.ToInt32();
                            //ITypeInfo refTypeInfo;
                            //info.GetRefTypeInfo(href, out refTypeInfo);

                            // todo: get type info?
                        }
                        if (valueType == VarEnum.VT_CARRAY || valueType == VarEnum.VT_ARRAY || valueType == VarEnum.VT_SAFEARRAY)
                        {
                            // todo: tell ParamArray arrays from normal arrays
                            isArray = true;
                        }

                        yield return new ParameterDeclaration(new QualifiedMemberName(typeQualifiedModuleName, paramName), memberDeclaration, asParamTypeName, isOptional, isByRef, isArray);
                    }
                }

                for (var fieldIndex = 0; fieldIndex < typeAttributes.cVars; fieldIndex++)
                {
                    IntPtr ppVarDesc;
                    info.GetVarDesc(fieldIndex, out ppVarDesc);

                    var varDesc = (VARDESC) Marshal.PtrToStructure(ppVarDesc, typeof (VARDESC));

                    var names = new string[255];
                    int namesArrayLength;
                    info.GetNames(varDesc.memid, names, 255, out namesArrayLength);

                    var fieldName = names[0];
                    var fieldValueType = (VarEnum)varDesc.elemdescVar.tdesc.vt;
                    var memberType = GetDeclarationType(varDesc, typeDeclarationType);

                    string asTypeName;
                    if (!TypeNames.TryGetValue(fieldValueType, out asTypeName))
                    {
                        asTypeName = TypeNames[VarEnum.VT_VARIANT];
                    }

                    yield return new Declaration(new QualifiedMemberName(typeQualifiedModuleName, fieldName), moduleDeclaration, moduleDeclaration, asTypeName, false, false, Accessibility.Global, memberType, null, Selection.Home);
                }
            }           
        }

        //private IEnumerable<string> GetImplementedInterfaceNames(TYPEATTR typeAttr, ITypeInfo info)
        //{
        //    for (var implIndex = 0; implIndex < typeAttr.cImplTypes; implIndex++)
        //    {
        //        int href;
        //        info.GetRefTypeOfImplType(implIndex, out href);

        //        ITypeInfo implTypeInfo;
        //        info.GetRefTypeInfo(href, out implTypeInfo);

        //        var implTypeName = GetTypeName(implTypeInfo);

        //        yield return implTypeName;
        //        //Debug.WriteLine(string.Format("\tImplements {0}", implTypeName));
        //    }
        //}

        private DeclarationType GetDeclarationType(ITypeLib typeLibrary, int i)
        {
            TYPEKIND typeKind;
            typeLibrary.GetTypeInfoType(i, out typeKind);

            DeclarationType typeDeclarationType = DeclarationType.Control; // todo: a better default
            if (typeKind == TYPEKIND.TKIND_ENUM)
            {
                typeDeclarationType = DeclarationType.Enumeration;
            }
            else if (typeKind == TYPEKIND.TKIND_COCLASS || typeKind == TYPEKIND.TKIND_INTERFACE ||
                     typeKind == TYPEKIND.TKIND_ALIAS || typeKind == TYPEKIND.TKIND_DISPATCH)
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
            return typeDeclarationType;
        }

        private DeclarationType GetDeclarationType(FUNCDESC funcDesc, VarEnum funcValueType)
        {
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
            return memberType;
        }

        private DeclarationType GetDeclarationType(VARDESC varDesc, DeclarationType typeDeclarationType)
        {
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
            return memberType;
        }
    }
}
