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
        IVerify Verify { get; }
        [DispId(2)]
        void AssignsByRef(string Parameter, object Value);
        [DispId(3)]
        void RaisesError(int Number = 0, string Description = "");
        [DispId(4)]
        bool PassThrough { get; set; }
        [DispId(5)]
        void Returns(object Value, int Invocation = FakesProvider.AllInvocations);
        [DispId(6)]
        void ReturnsWhen(string Parameter, object Argument, object Value, int Invocation = FakesProvider.AllInvocations);
    }
}
