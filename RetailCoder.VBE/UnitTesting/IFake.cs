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
        [Description("Gets an interface for verifying invocations performed during the test.")]
        IVerify Verify { get; }

        [DispId(2)]
        [Description("Configures the fake such as an invocation assigns the specified value to the specified ByRef argument.")]
        void AssignsByRef(string Parameter, object Value);

        [DispId(3)]
        [Description("Configures the fake such as an invocation raises the specified run-time eror.")]
        void RaisesError(int Number = 0, string Description = "");

        [DispId(4)]
        [Description("Gets/sets a value that determines whether execution is handled by Rubberduck.")]
        bool PassThrough { get; set; }

        [DispId(5)]
        [Description("Configures the fake such as the specified invocation returns the specified value.")]
        void Returns(object Value, int Invocation = FakesProvider.AllInvocations);

        [DispId(6)]
        [Description("Configures the fake such as the specified invocation returns the specified value given a specific parameter value.")]
        void ReturnsWhen(string Parameter, object Argument, object Value, int Invocation = FakesProvider.AllInvocations);
    }
}
