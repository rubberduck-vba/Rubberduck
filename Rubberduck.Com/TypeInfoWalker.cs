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
    public class TypeInfoWalker : TypeApiWalker<ITypeInfoVisitor>
    {
        private readonly ITypeLib _typeLib;
        private readonly ITypeLib2 _typeLib2;
        private readonly ITypeInfo _typeInfo;
        private readonly ITypeInfo2 _typeInfo2;
        private readonly int _index;

        private TypeInfoWalker(ITypeLib typeLib, int index, ITypeInfo typeInfo, IEnumerable<ITypeInfoVisitor> visitors)
        {
            _typeLib = typeLib;
            _typeLib2 = typeLib as ITypeLib2;
            _index = index;
            _typeInfo = typeInfo;
            _typeInfo2 = typeInfo as ITypeInfo2;
            Visitors = visitors;
        }

        public static void Accept(ITypeLib typeLib, int index, ITypeInfo typeInfo, ITypeInfoVisitor visitor) => Accept(typeLib, index, typeInfo, new[] { visitor });

        public static void Accept(ITypeLib typeLib, int index, ITypeInfo typeInfo, IEnumerable<ITypeInfoVisitor> visitors)
        {
            var walker = new TypeInfoWalker(typeLib, index, typeInfo, visitors);
            walker.WalkTypeInfo();
        }

        private void WalkTypeInfo()
        {
            var cImplTypes = 0;
            var cVars = 0;
            var cFuncs = 0;

            _typeLib.GetDocumentation(_index, out var name, out var docString, out var helpContext, out var helpFile);
            ExecuteVisit(v => v.VisitTypeDocumentation(name, docString, helpContext, helpFile));

            _typeInfo.UsingTypeAttr(attr => {
                cImplTypes = attr.cImplTypes;
                cVars = attr.cVars;
                cFuncs = attr.cFuncs;

                ExecuteVisit(v => v.VisitTypeAttr(attr));
            });

            if (_typeLib2 != null)
            {
                _typeLib2.GetDocumentation2(_index, out var helpString, out var helpStringContext, out var helpStringDll);
                ExecuteVisit(v => v.VisitTypeDocumentation2(helpString, helpStringContext, helpStringDll));
            }

            if (_typeInfo2 != null)
            {
                EnumerateCustomData(
                    _typeInfo2.GetAllCustData,
                    (guid, value) => ExecuteVisit(v => v.VisitTypeCustData(guid, value)));
            }

            EnumerateImplementedTypes(cImplTypes);
            EnumerateVariables(cVars);
            EnumerateFunctions(cFuncs);
        }

        private void EnumerateImplementedTypes(int cImplTypes)
        {
            for (var i = 0; i < cImplTypes; i++)
            {
                _typeInfo.GetRefTypeOfImplType(i, out var href);
                _typeInfo.GetRefTypeInfo(href, out var implementedTypeInfo);

                var implementedLibraryVisitors = new List<ITypeLibVisitor>();
                var implementedTypeVisitors = new List<ITypeInfoVisitor>();
                ExecuteVisit(v =>
                {
                    var result = v.VisitTypeImplementedTypeInfo(href, implementedTypeInfo);
                    switch (result)
                    {
                        case VisitDirectives.VisitLibrary:
                            implementedLibraryVisitors.AddRange(v.ProvideTypeLibVisitors().Where(x => !implementedLibraryVisitors.Contains(x)));
                            break;
                        case VisitDirectives.VisitType:
                            implementedTypeVisitors.Add(v);
                            break;
                    }
                });
                implementedTypeInfo.GetContainingTypeLib(out var implementedTypeLib, out var implementedIndex);
                if (implementedLibraryVisitors.Any())
                {
                    TypeLibWalker.Accept(implementedTypeLib, implementedLibraryVisitors);
                }
                if (implementedTypeVisitors.Any())
                {
                    Accept(implementedTypeLib, implementedIndex, implementedTypeInfo, implementedTypeVisitors);
                }

                if (_typeInfo2 != null)
                {
                    EnumerateCustomData(
                        ptr => _typeInfo2.GetAllImplTypeCustData(i, ptr),
                        (g, o) => ExecuteVisit(
                            v => v.VisitTypeImplTypeCustData(g, o)));
                }
            }
        }

        private void EnumerateVariables(int cVars)
        {
            var names = new string[1];
            for (var i = 0; i < cVars; i++)
            {
                _typeInfo.UsingVarDesc(i, varDesc =>
                {
                    _typeInfo.GetNames(varDesc.memid, names, 1, out var count);
                    object value = null;
                    if (varDesc.varkind == VARKIND.VAR_CONST && varDesc.desc.lpvarValue != IntPtr.Zero)
                    {
                        value = Marshal.GetObjectForNativeVariant(varDesc.desc.lpvarValue);
                    }

                    var internalType = GetTypeDesc(varDesc.elemdescVar.tdesc, _typeInfo);
                    var type = new VarDescData(names[0], varDesc, internalType, value);
                    ExecuteVisit(v => v.VisitTypeVarDesc(type));

                    if (_typeInfo2 != null)
                    {
                        EnumerateCustomData(
                            ptr => _typeInfo2.GetAllVarCustData(i, ptr),
                            (g, o) => ExecuteVisit(
                                v => v.VisitTypeVarCustData(g, o)));
                    }
                });
            }
        }

        private VarDescDataInternal GetTypeDesc(TYPEDESC typeDesc, ITypeInfo typeInfo)
        {
            var vt = (VarEnum)typeDesc.vt;
            TYPEDESC tdesc;

            if (vt == VarEnum.VT_PTR)
            {
                tdesc = Marshal.PtrToStructure<TYPEDESC>(typeDesc.lpValue);
                var result = GetTypeDesc(tdesc, typeInfo);
                result.IsReferencedType = true;
                return result;
            }
            else if (vt == VarEnum.VT_USERDEFINED)
            {
                int href;
                unchecked
                {
                    //The href is a long, but the size of lpValue depends on the platform, so truncate it after the lword.
                    href = (int)(typeDesc.lpValue.ToInt64() & 0xFFFFFFFF);
                }
                typeInfo.GetRefTypeInfo(href, out ITypeInfo refTypeInfo);
                var varDescType = new VarDescDataInternal();
                refTypeInfo.UsingTypeAttr(attr =>
                {
                    varDescType.IsReferencedType = ReferenceTypeKinds.Contains(attr.typekind);
                    varDescType.ReferencedTypeInfo = refTypeInfo;
                });
                return varDescType;
            }
            else if (vt == VarEnum.VT_SAFEARRAY || vt == VarEnum.VT_CARRAY || vt.HasFlag(VarEnum.VT_ARRAY))
            {
                tdesc = Marshal.PtrToStructure<TYPEDESC>(typeDesc.lpValue);
                var result = GetTypeDesc(tdesc, typeInfo);
                result.IsArray = true;
                return result;
            }

            return new VarDescDataInternal { VarEnum = vt };
        }

        private void EnumerateFunctions(int cFuncs)
        {
            for (var i = 0; i < cFuncs; i++)
            {
                _typeInfo.UsingFuncDesc(i, funcDesc =>
                {
                    var parameters = CollectParameters(i, funcDesc, _typeInfo);
                    _typeLib.GetDocumentation(funcDesc.memid, out var funcName, out var docString, out var helpContext, out var helpFile);
                    var data = new FuncDescData
                    (
                        i,
                        funcName,
                        docString,
                        helpContext,
                        helpFile,
                        funcDesc,
                        parameters
                    );

                    ExecuteVisit(v => v.VisitTypeFuncDesc(data));

                    if (_typeInfo2 != null)
                    {
                        EnumerateCustomData(
                            ptr => _typeInfo2.GetAllFuncCustData(i, ptr),
                            (g, o) => ExecuteVisit(v => v.VisitTypeFuncCustData(g, o)));
                    }
                });
            }
        }

        private IEnumerable<ParamDescData> CollectParameters(int index, FUNCDESC funcDesc, ITypeInfo typeInfo)
        {
            var parameters = new List<ParamDescData>();
            var names = new string[funcDesc.cParams + 1];
            typeInfo.GetNames(index, names, names.Length, out _);

            for (var i = 0; i < funcDesc.cParams; i++)
            {
                var paramPtr = funcDesc.lprgelemdescParam + (Marshal.SizeOf(typeof(ELEMDESC)) * i);
                var elemDesc = Marshal.PtrToStructure<ELEMDESC>(paramPtr);
                var (isByRef, isArray, vt, parameterTypeInfo, exception) = GetParameterType(elemDesc.tdesc, typeInfo);
                var customData = new Dictionary<Guid, object>();
                var ptrDefaultValue = elemDesc.desc.paramdesc.lpVarValue + Marshal.SizeOf(typeof(ulong));
                object defaultValue = GetParameterDefaultValue(ptrDefaultValue);

                if (typeInfo is ITypeInfo2 typeInfo2)
                {
                    EnumerateCustomData(
                        ptr => typeInfo2.GetAllParamCustData(index, i, ptr),
                        (g, v) => customData.Add(g, v));
                }

                var data = new ParamDescData(
                    i,
                    names[i + 1],
                    elemDesc,
                    isArray,
                    isByRef,
                    (i == funcDesc.cParams && funcDesc.cParamsOpt == -1),
                    elemDesc.desc.paramdesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FOPT),
                    vt,
                    defaultValue,
                    parameterTypeInfo,
                    exception,
                    customData
                );

                var implementedTypeInfo = parameterTypeInfo;
                var implementedLibraryVisitors = new List<ITypeLibVisitor>();
                var implementedTypeVisitors = new List<ITypeInfoVisitor>();
                ExecuteVisit(v =>
                {
                    var result = v.VisitTypeFuncParameter(data);
                    switch (result)
                    {
                        case VisitDirectives.VisitLibrary:
                            implementedLibraryVisitors.AddRange(v.ProvideTypeLibVisitors().Where(x => !implementedLibraryVisitors.Contains(x)));
                            break;
                        case VisitDirectives.VisitType:
                            implementedTypeVisitors.Add(v);
                            break;
                    }
                });
                implementedTypeInfo.GetContainingTypeLib(out var implementedTypeLib, out var implementedIndex);
                if (implementedLibraryVisitors.Any())
                {
                    TypeLibWalker.Accept(implementedTypeLib, implementedLibraryVisitors);
                }
                if (implementedTypeVisitors.Any())
                {
                    Accept(implementedTypeLib, implementedIndex, implementedTypeInfo, implementedTypeVisitors);
                }

                parameters.Add(data);
            }

            return parameters;
        }

        private (VarEnum vt, object value) GetParameterDefaultValue(IntPtr variant)
        {
            const ushort VT_TYPEMASK = 0xFFF;
            var members = Marshal.PtrToStructure<VARIANT>(variant);
            //AFAICT, this should always pass for automation types supported by VB(A). 
            Debug.Assert(!Convert.ToBoolean(~VT_TYPEMASK & (int)members.vt), "Non value-type will potentially leak a pointer.");

            var vt = (VarEnum)members.vt;
            var value = Marshal.GetObjectForNativeVariant(variant);

            if (value == null && vt == VarEnum.VT_BSTR)
            {
                value = string.Empty;
            }

            return (vt, value);
        }

        private (bool isByRef, bool isArray, VarEnum vt, ITypeInfo parameterTypeInfo, COMException exception) GetParameterType(TYPEDESC desc, ITypeInfo info)
        {
            var vt = (VarEnum)desc.vt;
            TYPEDESC tdesc;

            if (vt == VarEnum.VT_PTR)
            {
                tdesc = Marshal.PtrToStructure<TYPEDESC>(desc.lpValue);
                var (_, isArray, byRefvt, parameterTypeInfo, exception) = GetParameterType(tdesc, info);
                return (true, isArray, byRefvt, parameterTypeInfo, exception);
            }
            else if (vt == VarEnum.VT_USERDEFINED)
            {
                int href;
                unchecked
                {
                    href = (int)(desc.lpValue.ToInt64() & 0xFFFFFFFF);
                }

                try
                {
                    info.GetRefTypeInfo(href, out var refTypeInfo);
                    return (true, false, vt, refTypeInfo, null);
                }
                catch(COMException ex)
                {
                    return (true, false, vt, null, ex);
                }
            }
            else if (vt == VarEnum.VT_SAFEARRAY || vt == VarEnum.VT_CARRAY || vt.HasFlag(VarEnum.VT_ARRAY))
            {
                tdesc = Marshal.PtrToStructure<TYPEDESC>(desc.lpValue);
                var (isByRef, _, arrayVt, arrayTypeInfo, exception) = GetParameterType(tdesc, info);
                return (isByRef, true, arrayVt, arrayTypeInfo, exception);
            }
            else
            {
                return (false, false, vt, null, null);
            }
        }

        private static readonly HashSet<TYPEKIND> ReferenceTypeKinds = new HashSet<TYPEKIND>
        {
            TYPEKIND.TKIND_DISPATCH,
            TYPEKIND.TKIND_COCLASS,
            TYPEKIND.TKIND_INTERFACE
        };
    }
}
