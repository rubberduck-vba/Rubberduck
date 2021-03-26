using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

// ReSharper disable InconsistentNaming

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IComMockGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IComMock
    {
        [DispId(1)]
        [Description("Gets the mocked object.")]
        object Object { [return: MarshalAs(UnmanagedType.IDispatch)] get; }

        [DispId(2)]
        [Description("Gets the name of the loaded project defining the mocked interface.")]
        string Project { get; }

        [DispId(3)]
        [Description("Gets the programmatic name of the mocked interface.")]
        string ProgId { get; }

        [DispId(4)]
        [Description("Specifies a setup on the mocked type for a call to a method that does not return a value.")]
        void Setup(string Name, [MarshalAs(UnmanagedType.Struct)] object Args = null);

        [DispId(5)]
        [Description("Specifies a setup on the mocked type for a call to a value-returning method.")]
        void SetupWithReturns(string Name, [MarshalAs(UnmanagedType.Struct)] object Value, [Optional, MarshalAs(UnmanagedType.Struct)] object Args);

        [DispId(6)]
        [Description("Specifies a callback (use the AddressOf operator) to invoke when the method is called that receives the original invocation.")]
        void SetupWithCallback(string Name, [MarshalAs(UnmanagedType.FunctionPtr)] Action Callback, [Optional, MarshalAs(UnmanagedType.Struct)] object Args);

        [DispId(7)]
        [Description("Specifies a setup on the mocked type for a call to an object member of the specified object type.")]
        IComMock SetupChildMock(string Name, [Optional, MarshalAs(UnmanagedType.Struct)] object Args);

        [DispId(9)]
        [Description("Verifies that a specific invocation matching the given arguments was performed on the mock.")]
        void Verify(string Name, ITimes Times, [Optional, MarshalAs(UnmanagedType.Struct)] object Args);
    }
}