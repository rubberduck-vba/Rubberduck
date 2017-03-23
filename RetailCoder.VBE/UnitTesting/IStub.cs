using System.ComponentModel;
using System.Runtime.InteropServices;

// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.
// ReSharper disable InconsistentNaming

namespace Rubberduck.UnitTesting
{
    // IMPORTANT - C# doesn't support interface inheritance in its exported type libraries, so any members on this interface
    // should also be on the IFake interface with matching DispIds due to the inheritance of the concrete classes.

    [ComVisible(true)]
    [Guid(RubberduckGuid.IStubGuid)]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public interface IStub
    {
        [DispId(1)]
        IVerify Verify { get; }
        [DispId(2)]
        void AssignsByRef(string Parameter, object Value);
        [DispId(3)]
        void RaisesError(int Number = 0, string Description = "");
        [DispId(4)]
        bool PassThrough { get; set; }
    }
}
