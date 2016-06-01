using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using CALLCONV = System.Runtime.InteropServices.ComTypes.CALLCONV;
using FUNCFLAGS = System.Runtime.InteropServices.ComTypes.FUNCFLAGS;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using FUNCKIND = System.Runtime.InteropServices.ComTypes.FUNCKIND;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;

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
        private static extern Int32 LoadTypeLibEx(string strTypeLibName, REGKIND regKind, out ITypeLib TypeLib);

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
            {VarEnum.VT_CY, "Currency"},
            {VarEnum.VT_DECIMAL, "Currency"}, // best match?
            {VarEnum.VT_EMPTY, "Empty"},
            {VarEnum.VT_R4, "Single"},
            {VarEnum.VT_R8, "Double"},
        };

        private string GetTypeName(TYPEDESC desc, ITypeInfo info)
        {
            var vt = (VarEnum)desc.vt;
            TYPEDESC tdesc;

            switch (vt)
            {
                case VarEnum.VT_PTR:
                    tdesc = (TYPEDESC)Marshal.PtrToStructure(desc.lpValue, typeof(TYPEDESC));
                    return GetTypeName(tdesc, info);
                case VarEnum.VT_USERDEFINED:
                    int href;
                    unchecked
                    {
                        href = (int)(desc.lpValue.ToInt64() & 0xFFFFFFFF);
                    }
                    try
                    {
                        ITypeInfo refTypeInfo;
                        info.GetRefTypeInfo(href, out refTypeInfo);
                        return GetTypeName(refTypeInfo);
                    }
                    catch (Exception)
                    {
                        return "Object";
                    }
                case VarEnum.VT_CARRAY:
                    tdesc = (TYPEDESC)Marshal.PtrToStructure(desc.lpValue, typeof(TYPEDESC));
                    return GetTypeName(tdesc, info) + "()";
                default:
                    string result;
                    if (TypeNames.TryGetValue(vt, out result))
                    {
                        return result;
                    }
                    break;
            }
            return "Object";
        }

        private string GetTypeName(ITypeInfo info)
        {
            string typeName;
            string docString; // todo: put the docString to good use?
            int helpContext;
            string helpFile;
            info.GetDocumentation(-1, out typeName, out docString, out helpContext, out helpFile);

            return typeName;
        }

        public List<Declaration> GetDeclarationsForReference(Reference reference)
        {
            var output = new List<Declaration>();
            var projectName = reference.Name;
            var path = reference.FullPath;
            ITypeLib typeLibrary;
            // Failure to load might mean that it's a "normal" VBProject that will get parsed by us anyway.
            LoadTypeLibEx(path, REGKIND.REGKIND_NONE, out typeLibrary);
            if (typeLibrary == null)
            {
                return output;
            }
            var projectQualifiedModuleName = new QualifiedModuleName(projectName, path, projectName);
            var projectQualifiedMemberName = new QualifiedMemberName(projectQualifiedModuleName, projectName);
            var projectDeclaration = new ProjectDeclaration(projectQualifiedMemberName, projectName, isBuiltIn: true);
            output.Add(projectDeclaration);

            var typeCount = typeLibrary.GetTypeInfoCount();
            for (var i = 0; i < typeCount; i++)
            {
                ITypeInfo info;
                try
                {
                    typeLibrary.GetTypeInfo(i, out info);
                }
                catch (NullReferenceException)
                {
                    return output;
                }

                if (info == null)
                {
                    continue;
                }

                var typeName = GetTypeName(info);
                var typeDeclarationType = GetDeclarationType(typeLibrary, i);

                QualifiedModuleName typeQualifiedModuleName;
                QualifiedMemberName typeQualifiedMemberName;
                if (typeDeclarationType == DeclarationType.Enumeration || typeDeclarationType == DeclarationType.UserDefinedType)
                {
                    typeQualifiedModuleName = projectQualifiedModuleName;
                    typeQualifiedMemberName = new QualifiedMemberName(projectQualifiedModuleName, typeName);
                }
                else
                {
                    typeQualifiedModuleName = new QualifiedModuleName(projectName, path, typeName);
                    typeQualifiedMemberName = new QualifiedMemberName(typeQualifiedModuleName, typeName);
                }

                IntPtr typeAttributesPointer;
                info.GetTypeAttr(out typeAttributesPointer);

                var typeAttributes = (TYPEATTR)Marshal.PtrToStructure(typeAttributesPointer, typeof(TYPEATTR));

                var attributes = new Attributes();
                if (typeAttributes.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FPREDECLID))
                {
                    attributes.AddPredeclaredIdTypeAttribute();
                }

                Declaration moduleDeclaration;
                switch (typeDeclarationType)
                {
                    case DeclarationType.ProceduralModule:
                        moduleDeclaration = new ProceduralModuleDeclaration(typeQualifiedMemberName, projectDeclaration, typeName, true, new List<IAnnotation>(), attributes);
                        break;
                    case DeclarationType.ClassModule:
                        var module = new ClassModuleDeclaration(typeQualifiedMemberName, projectDeclaration, typeName, true, new List<IAnnotation>(), attributes);
                        var implements = GetImplementedInterfaceNames(typeAttributes, info);
                        foreach (var supertypeName in implements)
                        {
                            module.AddSupertype(supertypeName);
                        }
                        moduleDeclaration = module;
                        break;
                    default:
                        string pseudoModuleName = string.Format("_{0}", typeName);
                        var pseudoParentModule = new ProceduralModuleDeclaration(
                            new QualifiedMemberName(projectQualifiedModuleName, pseudoModuleName),
                            projectDeclaration,
                            pseudoModuleName,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                        // Enums don't define their own type but have a declared type of "Long".
                        if (typeDeclarationType == DeclarationType.Enumeration)
                        {
                            typeName = Tokens.Long;
                        }
                        // UDTs and ENUMs don't seem to have a module parent that's why we add a "fake" module
                        // so that the rest of the application can treat it normally.
                        moduleDeclaration = new Declaration(
                            typeQualifiedMemberName,
                            pseudoParentModule,
                            pseudoParentModule, 
                            typeName,
                            null,
                            false, 
                            false, 
                            Accessibility.Global,
                            typeDeclarationType,
                            null, 
                            Selection.Home,
                            false,
                            null,
                            true, 
                            null, 
                            attributes);
                        break;
                }

                output.Add(moduleDeclaration);

                for (var memberIndex = 0; memberIndex < typeAttributes.cFuncs; memberIndex++)
                {
                    FUNCDESC memberDescriptor;
                    string[] memberNames;
                    var memberDeclaration = CreateMemberDeclaration(out memberDescriptor, typeAttributes.typekind, info, memberIndex, typeQualifiedModuleName, moduleDeclaration, out memberNames);
                    if (memberDeclaration == null)
                    {
                        continue;
                    }
                    if (moduleDeclaration.DeclarationType == DeclarationType.ClassModule && memberDeclaration is ICanBeDefaultMember && ((ICanBeDefaultMember)memberDeclaration).IsDefaultMember)
                    {
                        ((ClassModuleDeclaration)moduleDeclaration).DefaultMember = memberDeclaration;
                    }
                    output.Add(memberDeclaration);

                    var parameterCount = memberDescriptor.cParams - 1;
                    for (var paramIndex = 0; paramIndex < parameterCount; paramIndex++)
                    {
                        var parameter = CreateParameterDeclaration(memberNames, paramIndex, memberDescriptor, typeQualifiedModuleName, memberDeclaration, info);
                        var declaration = memberDeclaration as IDeclarationWithParameter;
                        if (declaration != null)
                        {
                            declaration.AddParameter(parameter);
                        }
                        output.Add(parameter);
                    }
                }

                for (var fieldIndex = 0; fieldIndex < typeAttributes.cVars; fieldIndex++)
                {
                    output.Add(CreateFieldDeclaration(info, fieldIndex, typeDeclarationType, typeQualifiedModuleName, moduleDeclaration));
                }
            }
            return output;
        }

        private Declaration CreateMemberDeclaration(out FUNCDESC memberDescriptor, TYPEKIND typeKind, ITypeInfo info, int memberIndex,
            QualifiedModuleName typeQualifiedModuleName, Declaration moduleDeclaration, out string[] memberNames)
        {
            IntPtr memberDescriptorPointer;
            info.GetFuncDesc(memberIndex, out memberDescriptorPointer);
            memberDescriptor = (FUNCDESC)Marshal.PtrToStructure(memberDescriptorPointer, typeof(FUNCDESC));

            if (memberDescriptor.callconv != CALLCONV.CC_STDCALL)
            {
                memberDescriptor = new FUNCDESC();
                memberNames = new string[] { };
                return null;
            }

            memberNames = new string[255];
            int namesArrayLength;
            info.GetNames(memberDescriptor.memid, memberNames, 255, out namesArrayLength);

            var memberName = memberNames[0];
            var funcValueType = (VarEnum)memberDescriptor.elemdescFunc.tdesc.vt;
            var memberDeclarationType = GetDeclarationType(memberDescriptor, funcValueType, typeKind);

            var asTypeName = string.Empty;
            if (memberDeclarationType != DeclarationType.Procedure)
            {
                asTypeName = GetTypeName(memberDescriptor.elemdescFunc.tdesc, info);
            }
            var attributes = new Attributes();
            if (memberName == "_NewEnum" && ((FUNCFLAGS)memberDescriptor.wFuncFlags).HasFlag(FUNCFLAGS.FUNCFLAG_FNONBROWSABLE))
            {
                attributes.AddEnumeratorMemberAttribute(memberName);
            }
            else if (memberDescriptor.memid == 0)
            {
                attributes.AddDefaultMemberAttribute(memberName);
            }
            else if (((FUNCFLAGS)memberDescriptor.wFuncFlags).HasFlag(FUNCFLAGS.FUNCFLAG_FHIDDEN))
            {
                attributes.AddHiddenMemberAttribute(memberName);
            }

            switch (memberDeclarationType)
            {
                case DeclarationType.Procedure:
                    return new SubroutineDeclaration(
                        new QualifiedMemberName(typeQualifiedModuleName, memberName),
                        moduleDeclaration,
                        moduleDeclaration,
                        asTypeName,
                        Accessibility.Global,
                        null,
                        Selection.Home,
                        true,
                        null,
                        attributes);
                case DeclarationType.Function:
                    return new FunctionDeclaration(
                        new QualifiedMemberName(typeQualifiedModuleName, memberName),
                        moduleDeclaration,
                        moduleDeclaration,
                        asTypeName,
                        null,
                        null,
                        Accessibility.Global,
                        null,
                        Selection.Home,
                        // TODO: how to find out if it's an array?
                        false,
                        true,
                        null,
                        attributes);
                case DeclarationType.PropertyGet:
                    return new PropertyGetDeclaration(
                        new QualifiedMemberName(typeQualifiedModuleName, memberName),
                        moduleDeclaration,
                        moduleDeclaration,
                        asTypeName,
                        null,
                        null,
                        Accessibility.Global,
                        null,
                        Selection.Home,
                        // TODO: how to find out if it's an array?
                        false,
                        true,
                        null,
                        attributes);
                case DeclarationType.PropertySet:
                    return new PropertySetDeclaration(
                        new QualifiedMemberName(typeQualifiedModuleName, memberName),
                        moduleDeclaration,
                        moduleDeclaration,
                        asTypeName,
                        Accessibility.Global,
                        null,
                        Selection.Home,
                        true,
                        null,
                        attributes);
                case DeclarationType.PropertyLet:
                    return new PropertyLetDeclaration(
                        new QualifiedMemberName(typeQualifiedModuleName, memberName),
                        moduleDeclaration,
                        moduleDeclaration,
                        asTypeName,
                        Accessibility.Global,
                        null,
                        Selection.Home,
                        true,
                        null,
                        attributes);
                default:
                    return new Declaration(
                        new QualifiedMemberName(typeQualifiedModuleName, memberName),
                        moduleDeclaration,
                        moduleDeclaration,
                        asTypeName,
                        null,
                        false,
                        false,
                        Accessibility.Global,
                        memberDeclarationType,
                        null,
                        Selection.Home,
                        false,
                        null,
                        true,
                        null,
                        attributes);
            }
        }

        private Declaration CreateFieldDeclaration(ITypeInfo info, int fieldIndex, DeclarationType typeDeclarationType,
            QualifiedModuleName typeQualifiedModuleName, Declaration moduleDeclaration)
        {
            IntPtr ppVarDesc;
            info.GetVarDesc(fieldIndex, out ppVarDesc);

            var varDesc = (VARDESC)Marshal.PtrToStructure(ppVarDesc, typeof(VARDESC));

            var names = new string[255];
            int namesArrayLength;
            info.GetNames(varDesc.memid, names, 255, out namesArrayLength);

            var fieldName = names[0];
            var memberType = GetDeclarationType(varDesc, typeDeclarationType);

            var asTypeName = GetTypeName(varDesc.elemdescVar.tdesc, info);            

            return new Declaration(new QualifiedMemberName(typeQualifiedModuleName, fieldName),
                moduleDeclaration, moduleDeclaration, asTypeName, null, false, false, Accessibility.Global, memberType, null,
                Selection.Home, false, null);
        }

        private ParameterDeclaration CreateParameterDeclaration(IReadOnlyList<string> memberNames, int paramIndex,
            FUNCDESC memberDescriptor, QualifiedModuleName typeQualifiedModuleName, Declaration memberDeclaration, ITypeInfo info)
        {
            var paramName = memberNames[paramIndex + 1];

            var paramPointer = new IntPtr(memberDescriptor.lprgelemdescParam.ToInt64() + Marshal.SizeOf(typeof(ELEMDESC)) * paramIndex);
            var elementDesc = (ELEMDESC)Marshal.PtrToStructure(paramPointer, typeof(ELEMDESC));
            var isOptional = elementDesc.desc.paramdesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FOPT);

            var isByRef = elementDesc.desc.paramdesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FOUT);
            var isArray = false;
            var paramDesc = elementDesc.tdesc;
            var valueType = (VarEnum)paramDesc.vt;
            if (valueType == VarEnum.VT_CARRAY || valueType == VarEnum.VT_ARRAY || valueType == VarEnum.VT_SAFEARRAY)
            {
                // todo: tell ParamArray arrays from normal arrays
                isArray = true;
            }

            var asParamTypeName = GetTypeName(paramDesc, info);

            return new ParameterDeclaration(new QualifiedMemberName(typeQualifiedModuleName, paramName), memberDeclaration, asParamTypeName, null, null, isOptional, isByRef, isArray);
        }

        private IEnumerable<string> GetImplementedInterfaceNames(TYPEATTR typeAttr, ITypeInfo info)
        {
            var output = new List<string>();
            for (var implIndex = 0; implIndex < typeAttr.cImplTypes; implIndex++)
            {
                int href;
                info.GetRefTypeOfImplType(implIndex, out href);

                ITypeInfo implTypeInfo;
                info.GetRefTypeInfo(href, out implTypeInfo);

                var implTypeName = GetTypeName(implTypeInfo);
                if (implTypeName != "IDispatch" && implTypeName != "IUnknown")
                {
                    // skip IDispatch.. just about everything implements it and RD doesn't need to care about it; don't care about IUnknown either
                    output.Add(implTypeName);
                }
                //Debug.WriteLine(string.Format("\tImplements {0}", implTypeName));
            }
            return output;
        }

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
                typeDeclarationType = DeclarationType.ClassModule;
            }
            else if (typeKind == TYPEKIND.TKIND_RECORD)
            {
                typeDeclarationType = DeclarationType.UserDefinedType;
            }
            else if (typeKind == TYPEKIND.TKIND_MODULE)
            {
                typeDeclarationType = DeclarationType.ProceduralModule;
            }
            return typeDeclarationType;
        }

        private DeclarationType GetDeclarationType(FUNCDESC funcDesc, VarEnum funcValueType, TYPEKIND typekind)
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
            else if (funcDesc.funckind == FUNCKIND.FUNC_PUREVIRTUAL && typekind == TYPEKIND.TKIND_COCLASS)
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
