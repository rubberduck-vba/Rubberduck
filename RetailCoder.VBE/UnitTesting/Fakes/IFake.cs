using System.ComponentModel;
using System.Runtime.InteropServices;

// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.
// ReSharper disable InconsistentNaming

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    [Guid(RubberduckGuid.IFakeGuid)]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public interface IFake
    {
        [DispId(1)]
        void Returns(object Value);
        [DispId(2)]
        void AssignsByRef(string Parameter, object Value);
        [DispId(3)]
        void RaisesError(int Number = 0, string Description = "");
        [DispId(4)]
        IVerify Verify { get; }
    }
}
