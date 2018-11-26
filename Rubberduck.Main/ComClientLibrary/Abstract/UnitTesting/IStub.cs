using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.
// ReSharper disable InconsistentNaming

namespace Rubberduck.UnitTesting
{
    // IMPORTANT - C# doesn't support interface inheritance in its exported type libraries, so any members on this interface
    // should also be on the IFake interface with matching DispIds due to the inheritance of the concrete classes.
    [
        ComVisible(true),
        Guid(RubberduckGuid.IStubGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public interface IStub
    {
        [DispId(1)]
        [Description("Gets an interface for verifying invocations performed during the test.")]
        IVerify Verify { get; }

        [DispId(2)]
        [Description("Configures the stub such as an invocation assigns the specified value to the specified ByRef argument.")]
        void AssignsByRef(string Parameter, object Value);

        [DispId(3)]
        [Description("Configures the stub such as an invocation raises the specified run-time eror.")]
        void RaisesError(int Number = 0, string Description = "");

        [DispId(4)]
        [Description("Gets/sets a value that determines whether execution is handled by Rubberduck.")]
        bool PassThrough { get; set; }
    }
}
