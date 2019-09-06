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
        void SetupWithReturns(string Name, [MarshalAs(UnmanagedType.Struct)] object Value, [Optional, MarshalAs(UnmanagedType.Struct)] object Args);

        [DispId(3)]
        void SetupWithCallback(string Name, [MarshalAs(UnmanagedType.FunctionPtr)] Action Callback, [Optional, MarshalAs(UnmanagedType.Struct)] object Args);

        [DispId(4)]
        IComMock SetupChildMock(string Name, [Optional, MarshalAs(UnmanagedType.Struct)] object Args);

        [DispId(5)]
        string Project { get; }

        [DispId(6)]
        string ProgId { get; }
    }
}