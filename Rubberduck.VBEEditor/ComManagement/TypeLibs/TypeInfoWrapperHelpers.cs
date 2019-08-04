using System;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Static helpers here for working with <see cref="ITypeInfoWrapper"/>
    /// </summary>
    internal static class TypeInfoWrapperHelpers
    {
        /// <summary>
        /// Gets the control <see cref="ComTypes.ITypeInfo"/> by looking for the
        /// corresponding getter on the form interface and returning its retval type
        /// </summary>
        /// <param name="controlName">the name of the control</param>
        /// <returns>
        /// <see cref="ITypeInfoWrapper"/> representing the type of control,
        /// typically the coclass, but this is host dependent
        /// </returns>
        public static ITypeInfoWrapper GetControlTypeFromInterface(ITypeInfoWrapper rootInterface, string controlName)
        {
            // TODO should encapsulate handling of raw datatypes
            foreach (var func in rootInterface.Funcs)
            {
                using (func)
                {
                    // Controls are exposed as getters on the interface.
                    //     can either be    ControlType* get_ControlName()       
                    //     or               HRESULT get_ControlName(ControlType** Out) 

                    if ((func.Name == controlName) &&
                        (func.ProcKind == PROCKIND.PROCKIND_GET) &&
                        (func.ParamCount == 0) &&
                        (func.FuncDesc.elemdescFunc.tdesc.vt == (short)VarEnum.VT_PTR))
                    {
                        var retValElement = StructHelper.ReadStructureUnsafe<ComTypes.ELEMDESC>(func.FuncDesc.elemdescFunc.tdesc.lpValue);
                        if (retValElement.tdesc.vt == (short)VarEnum.VT_USERDEFINED)
                        {
                            var hr = rootInterface.GetSafeRefTypeInfo((int)retValElement.tdesc.lpValue, out var retVal);
                            if (ComHelper.HRESULT_FAILED(hr)) throw RdMarshal.GetExceptionForHR(hr);
                            return retVal;
                        }
                    }
                    else if ((func.Name == controlName) &&
                        (func.ProcKind == PROCKIND.PROCKIND_GET) &&
                        (func.ParamCount == 1) &&
                        (func.FuncDesc.elemdescFunc.tdesc.vt == (short)VarEnum.VT_HRESULT))
                    {
                        // Get details of the first argument
                        var retValElementOuterPtr = StructHelper.ReadStructureUnsafe<ComTypes.ELEMDESC>(func.FuncDesc.lprgelemdescParam);
                        if (retValElementOuterPtr.tdesc.vt == (short)VarEnum.VT_PTR)
                        {
                            var retValElementInnerPtr = StructHelper.ReadStructureUnsafe<ComTypes.ELEMDESC>(retValElementOuterPtr.tdesc.lpValue);
                            if (retValElementInnerPtr.tdesc.vt == (short)VarEnum.VT_PTR)
                            {
                                var retValElement = StructHelper.ReadStructureUnsafe<ComTypes.ELEMDESC>(retValElementInnerPtr.tdesc.lpValue);

                                if (retValElement.tdesc.vt == (short)VarEnum.VT_USERDEFINED)
                                {
                                    var hr = rootInterface.GetSafeRefTypeInfo((int)retValElement.tdesc.lpValue, out var retVal);
                                    if (ComHelper.HRESULT_FAILED(hr)) throw RdMarshal.GetExceptionForHR(hr);
                                    return retVal;
                                }
                            }
                        }
                    }
                }
            }

            throw new ArgumentException($"TypeInfoWrapper::GetControlType failed. '{controlName}' control not found.");
        }
    }
}
