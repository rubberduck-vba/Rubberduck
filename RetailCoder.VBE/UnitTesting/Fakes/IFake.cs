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
        void Returns(object Value);
        void AssignsByRef(string Parameter, object Value);
        void RaisesError(int Number = 0, string Description = "");
        IVerify Verify { get; }
    }
}
