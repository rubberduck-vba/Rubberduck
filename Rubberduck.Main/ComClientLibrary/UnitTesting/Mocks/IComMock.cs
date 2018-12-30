using System;
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
        object Object { [return: MarshalAs(UnmanagedType.IDispatch)] get; }

        [DispId(2)]
        void SetupWithReturns(string Name, object Value, [Optional, MarshalAs(UnmanagedType.Struct)] object Args);

        [DispId(3)]
        void SetupWithCallback(string Name, [MarshalAs(UnmanagedType.FunctionPtr)] Action Callback, [Optional, MarshalAs(UnmanagedType.Struct)] object Args);
    }
}