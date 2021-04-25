using System;

namespace Rubberduck.Com
{
    /// <summary>
    /// There are several constants that are used within the type library APIs; this class helps
    /// encapsulates different constants for easy discovery. Each group of related constants should
    /// be in a nested class to allow us to use syntax like <see cref="Iids.IID_DISPATCH"/>
    /// to make it easier to locate the constant when programming against the API. 
    /// </summary>
    public static class WellKnown
    {
        public static class Iids
        {
            public static readonly Guid IID_UNKNOWN = new Guid("00000000-0000-0000-C000-000000000046");
            public static readonly Guid IID_DISPATCH = new Guid("00020400-0000-0000-C000-000000000046");
        }

        /// <summary>
        /// MS-OAUT Section 2.2.32.1
        /// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/cb9d0131-c6bd-463d-9c40-7264856a10c5
        /// Also see:
        /// https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/dispid-constants
        /// The lower DISPIDs constants are not used by all clients. For example, the <see cref="DISPID_CONSTRUCTOR"/> 
        /// and <see cref="DISPID_DESTRUCTOR"/> are used as part of DCOM but not normally within Automation. 
        /// </summary>
        public static class DispIds
        {
            public const int DISPID_VALUE = 0;
            public const int DISPID_UNKNOWN = -1;
            public const int DISPID_PROPERTYPUT = -3;
            public const int DISPID_NEWENUM = -4;
            public const int DISPID_EVALUATE = -5;
            public const int DISPID_CONSTRUCTOR = -6;
            public const int DISPID_DESTRUCTOR = -7;
            public const int DISPID_COLLECT = -8;
        }

        /// <summary>
        /// MS-OAUT Section 2.2.35.1
        /// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/5fbb4851-25f6-45ef-9f83-e9dd633e1e00
        /// </summary>
        public static class MemberIds
        {
            public const int MEMBERID_NIL = -1;
            public const int MEMBERID_DEFAULTINST = -2;
        }

        /// <summary>
        /// Used with <see cref="ICreateTypeInfo.AddImplType(uint, uint)"/>'s first parameter, based on the 
        /// documentation referred here:
        /// https://docs.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-icreatetypeinfo-addimpltype
        /// </summary>
        public static class ImplIndexes
        {
            /// <summary>
            /// The implementation must be a IUnknown-derived interface for use in a dual implementation
            /// </summary>
            public const int DualUnknown = -1;
            /// <summary>
            /// The base interface for which the current <see cref="ICreateTypeInfo"/> derives from. Normally, 
            /// this is either the <code>IUnknown</code> or <code>IDispatch"</code> interface. It may derive
            /// from another interface as long the root is one of either. The referenced <see cref="ITypeInfo"/>
            /// must be of <see cref="TYPEKIND.TKIND_INTERFACE"/>.
            /// </summary>
            public const int BaseInterface = 0;
            /// <summary>
            /// The IUnknown implementation of the interface for which the dispatch interface must be
            /// based on. 
            /// </summary>
            public const int DispatchInterface = 1;
        }
    }
}
