﻿using System;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.Parsing
{
    /// <summary>
    /// https://www.add-in-express.com/creating-addins-blog/2011/12/20/type-name-system-comobject/
    /// </summary>
    public class ComHelper
    {
        /// <summary>
        /// Returns a string value representing the type name of the specified COM object.
        /// </summary>
        /// <param name="comObj">A COM object the type name of which to return.</param>
        /// <returns>A string containing the type name.</returns>
        public static string GetTypeName(object comObj)
        {

            if (comObj == null)
                return string.Empty;

            if (!Marshal.IsComObject(comObj))
                //The specified object is not a COM object
                return string.Empty;

            if (!(comObj is IDispatch dispatch))
                //The specified COM object doesn't support getting type information
                return string.Empty;

            ComTypes.ITypeInfo typeInfo = null;
            try
            {
                try
                {
                    // obtain the ITypeInfo interface from the object
                    dispatch.GetTypeInfo(0, 0, out typeInfo);
                }
                catch (Exception)
                {
                    //Cannot get the ITypeInfo interface for the specified COM object
                    return string.Empty;
                }

                var typeName = "";

                try
                {
                    //retrieves the documentation string for the specified type description 
                    typeInfo.GetDocumentation(-1, out typeName, out _, out _, out _);
                }
                catch (Exception)
                {
                    // Cannot extract ITypeInfo information
                    return string.Empty;
                }
                return typeName;
            }
            catch (Exception)
            {
                // Unexpected error
                return string.Empty;
            }
            finally
            {
                if (typeInfo != null)
                {
                    Marshal.ReleaseComObject(typeInfo);
                }
            }
        }
    }

    /// <summary>
    /// Exposes objects, methods and properties to programming tools and other
    /// applications that support Automation.
    /// </summary>
    [ComImport()]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatch
    {
        [PreserveSig]
        int GetTypeInfoCount(out int Count);

        [PreserveSig]
        int GetTypeInfo(
            [MarshalAs(UnmanagedType.U4)] int iTInfo,
            [MarshalAs(UnmanagedType.U4)] int lcid,
            out ComTypes.ITypeInfo typeInfo);

        [PreserveSig]
        int GetIDsOfNames(
            ref Guid riid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)]
			string[] rgsNames,
            int cNames,
            int lcid,
            [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);

        [PreserveSig]
        int Invoke(
            int dispIdMember,
            ref Guid riid,
            uint lcid,
            ushort wFlags,
            ref ComTypes.DISPPARAMS pDispParams,
            out object pVarResult,
            ref ComTypes.EXCEPINFO pExcepInfo,
            IntPtr[] pArgErr);
    }
}
