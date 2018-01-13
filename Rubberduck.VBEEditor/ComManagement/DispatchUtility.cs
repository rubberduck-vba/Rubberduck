using System;
using System.Runtime.InteropServices;
using System.Reflection;

namespace Rubberduck.VBEditor.ComManagement
{
    /// <summary>
    /// Provides helper methods for working with COM IDispatch objects that have a registered type library.
    /// </summary>
    public static class DispatchUtility
    {
        private const int S_OK = 0; //From WinError.h
        private const int LOCALE_SYSTEM_DEFAULT = 2 << 10; //From WinNT.h == 2048 == 0x800

        /// <summary>
        /// Gets whether the specified object implements IDispatch.
        /// </summary>
        /// <param name="obj">An object to check.</param>
        /// <returns>True if the object implements IDispatch.  False otherwise.</returns>
        public static bool ImplementsIDispatch(object obj)
        {
            return obj is IDispatchInfo;
        }

        /// <summary>
        /// Gets a Type that can be used with reflection.
        /// </summary>
        /// <param name="obj">An object that implements IDispatch.</param>
        /// <param name="throwIfNotFound">Whether an exception should be thrown if a Type can't be obtained.</param>
        /// <returns>A .NET Type that can be used with reflection.</returns>
        /// <exception cref="InvalidCastException">If <paramref name="obj"/> doesn't implement IDispatch.</exception>
        public static Type GetType(object obj, bool throwIfNotFound)
        {
            RequireReference(obj, "obj");
            return GetType((IDispatchInfo)obj, throwIfNotFound);
        }

        /// <summary>
        /// Tries to get the DISPID for the requested member name.
        /// </summary>
        /// <param name="obj">An object that implements IDispatch.</param>
        /// <param name="name">The name of a member to lookup.</param>
        /// <param name="dispId">If the method returns true, this holds the DISPID on output.
        /// If the method returns false, this value should be ignored.</param>
        /// <returns>True if the member was found and resolved to a DISPID.  False otherwise.</returns>
        /// <exception cref="InvalidCastException">If <paramref name="obj"/> doesn't implement IDispatch.</exception>
        public static bool TryGetDispId(object obj, string name, out int dispId)
        {
            RequireReference(obj, "obj");
            return TryGetDispId((IDispatchInfo)obj, name, out dispId);
        }

        /// <summary>
        /// Invokes a member by DISPID.
        /// </summary>
        /// <param name="obj">An object that implements IDispatch.</param>
        /// <param name="dispId">The DISPID of a member.  This can be obtained using
        /// <see cref="TryGetDispId(object, string, out int)"/>.</param>
        /// <param name="args">The arguments to pass to the member.</param>
        /// <returns>The member's return value.</returns>
        /// <remarks>
        /// This can invoke a method or a property get accessor.
        /// </remarks>
        public static object Invoke(object obj, int dispId, object[] args, BindingFlags flags)
        {
            var memberName = "[DispId=" + dispId + "]";
            return Invoke(obj, memberName, args, flags);
        }

        /// <summary>
        /// Invokes a member by name.
        /// </summary>
        /// <param name="obj">An object.</param>
        /// <param name="memberName">The name of the member to invoke.</param>
        /// <param name="args">The arguments to pass to the member.</param>
        /// <returns>The member's return value.</returns>
        /// <remarks>
        /// This can invoke a method or a property get accessor.
        /// </remarks>
        public static object Invoke(object obj, string memberName, object[] args, BindingFlags flags)
        {
            RequireReference(obj, "obj");
            Type type = obj.GetType();
            return type.InvokeMember(memberName, flags,
                null, obj, args, null);
        }
        
        /// <summary>
        /// Requires that the value is non-null.
        /// </summary>
        /// <typeparam name="T">The type of the value.</typeparam>
        /// <param name="value">The value to check.</param>
        /// <param name="name">The name of the value.</param>
        private static void RequireReference<T>(T value, string name) where T : class
        {
            if (value == null)
            {
                throw new ArgumentNullException(name);
            }
        }

        /// <summary>
        /// Gets a Type that can be used with reflection.
        /// </summary>
        /// <param name="dispatch">An object that implements IDispatch.</param>
        /// <param name="throwIfNotFound">Whether an exception should be thrown if a Type can't be obtained.</param>
        /// <returns>A .NET Type that can be used with reflection.</returns>
        private static Type GetType(IDispatchInfo dispatch, bool throwIfNotFound)
        {
            RequireReference(dispatch, "dispatch");

            Type result = null;
            var hr = dispatch.GetTypeInfoCount(out var typeInfoCount);
            if (hr == S_OK && typeInfoCount > 0)
            {
                // Type info isn't usually culture-aware for IDispatch, so we might as well pass
                // the default locale instead of looking up the current thread's LCID each time
                // (via CultureInfo.CurrentCulture.LCID).
                dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, out result);
            }

            if (result == null && throwIfNotFound)
            {
                // If the GetTypeInfoCount called failed, throw an exception for that.
                Marshal.ThrowExceptionForHR(hr);

                // Otherwise, throw the same exception that Type.GetType would throw.
                throw new TypeLoadException();
            }

            return result;
        }

        /// <summary>
        /// Tries to get the DISPID for the requested member name.
        /// </summary>
        /// <param name="dispatch">An object that implements IDispatch.</param>
        /// <param name="name">The name of a member to lookup.</param>
        /// <param name="dispId">If the method returns true, this holds the DISPID on output.
        /// If the method returns false, this value should be ignored.</param>
        /// <returns>True if the member was found and resolved to a DISPID.  False otherwise.</returns>
        private static bool TryGetDispId(IDispatchInfo dispatch, string name, out int dispId)
        {
            RequireReference(dispatch, "dispatch");
            RequireReference(name, "name");

            var result = false;

            // Members names aren't usually culture-aware for IDispatch, so we might as well
            // pass the default locale instead of looking up the current thread's LCID each time
            // (via CultureInfo.CurrentCulture.LCID).
            var iidNull = Guid.Empty;
            var hr = dispatch.GetDispId(ref iidNull, ref name, 1, LOCALE_SYSTEM_DEFAULT, out dispId);

            const int DISP_E_UNKNOWNNAME = unchecked((int)0x80020006); //From WinError.h
            const int DISPID_UNKNOWN = -1; //From OAIdl.idl
            switch (hr)
            {
                case S_OK:
                    result = true;
                    break;
                case DISP_E_UNKNOWNNAME when dispId == DISPID_UNKNOWN:
                    // This is the only supported "error" case because it means IDispatch
                    // is saying it doesn't know the member we asked about.
                    result = false;
                    break;
                default:
                    Marshal.ThrowExceptionForHR(hr);
                    break;
            }

            return result;
        }
        
        /// <summary>
        /// A partial declaration of IDispatch used to lookup Type information and DISPIDs.
        /// </summary>
        /// <remarks>
        /// This interface only declares the first three methods of IDispatch.  It omits the
        /// fourth method (Invoke) because there are already plenty of ways to do dynamic
        /// invocation in .NET.  But the first three methods provide dynamic type metadata
        /// discovery, which .NET doesn't provide normally if you have a System.__ComObject
        /// RCW instead of a strongly-typed RCW.
        /// <para/>
        /// Note: The original declaration of IDispatch is in OAIdl.idl.
        /// </remarks>
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("00020400-0000-0000-C000-000000000046")]
        private interface IDispatchInfo
        {
            /// <summary>
            /// Gets the number of Types that the object provides (0 or 1).
            /// </summary>
            /// <param name="typeInfoCount">Returns 0 or 1 for the number of Types provided by <see cref="GetTypeInfo"/>.</param>
            /// <remarks>
            /// http://msdn.microsoft.com/en-us/library/da876d53-cb8a-465c-a43e-c0eb272e2a12(VS.85)
            /// </remarks>
            [PreserveSig]
            int GetTypeInfoCount(out int typeInfoCount);

            /// <summary>
            /// Gets the Type information for an object if <see cref="GetTypeInfoCount"/> returned 1.
            /// </summary>
            /// <param name="typeInfoIndex">Must be 0.</param>
            /// <param name="lcid">Typically, LOCALE_SYSTEM_DEFAULT (2048).</param>
            /// <param name="typeInfo">Returns the object's Type information.</param>
            /// <remarks>
            /// http://msdn.microsoft.com/en-us/library/cc1ec9aa-6c40-4e70-819c-a7c6dd6b8c99(VS.85)
            /// </remarks>
            void GetTypeInfo(int typeInfoIndex, int lcid, [MarshalAs(UnmanagedType.CustomMarshaler,
                MarshalTypeRef = typeof(System.Runtime.InteropServices.CustomMarshalers.TypeToTypeInfoMarshaler))] out Type typeInfo);

            /// <summary>
            /// Gets the DISPID of the specified member name.
            /// </summary>
            /// <param name="riid">Must be IID_NULL.  Pass a copy of Guid.Empty.</param>
            /// <param name="name">The name of the member to look up.</param>
            /// <param name="nameCount">Must be 1.</param>
            /// <param name="lcid">Typically, LOCALE_SYSTEM_DEFAULT (2048).</param>
            /// <param name="dispId">If a member with the requested <paramref name="name"/>
            /// is found, this returns its DISPID and the method's return value is 0.
            /// If the method returns a non-zero value, then this parameter's output value is
            /// undefined.</param>
            /// <returns>Zero for success. Non-zero for failure.</returns>
            /// <remarks>
            /// http://msdn.microsoft.com/en-us/library/6f6cf233-3481-436e-8d6a-51f93bf91619(VS.85)
            /// </remarks>
            [PreserveSig]
            int GetDispId(ref Guid riid, ref string name, int nameCount, int lcid, out int dispId);

            // NOTE: The real IDispatch also has an Invoke method next, but we don't need it.
            // We can invoke methods using .NET's Type.InvokeMember method with the special
            // [DISPID=n] syntax for member "names", or we can get a .NET Type using GetTypeInfo
            // and invoke methods on that through reflection.
            // Type.InvokeMember: http://msdn.microsoft.com/en-us/library/de3dhzwy.aspx
        }
    }
}
