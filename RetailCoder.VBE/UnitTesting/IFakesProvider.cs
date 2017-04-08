using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    [Guid(RubberduckGuid.IFakesProviderGuid)]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public interface IFakesProvider
    {
        [DispId(1)]
        [Description("Configures VBA.Interactions.MsgBox calls.")]
        IFake MsgBox { get; }

        [DispId(2)]
        [Description("Configures VBA.Interactions.InputBox calls.")]
        IFake InputBox { get; }

        [DispId(3)]
        [Description("Configures VBA.Interaction.Beep calls.")]
        IStub Beep { get; }

        [DispId(4)]
        [Description("Configures VBA.Interaction.Environ calls.")]
        IFake Environ { get; }

        [DispId(5)]
        [Description("Configures VBA.DateTime.Timer calls.")]
        IFake Timer { get; }

        [DispId(6)]
        [Description("Configures VBA.Interaction.DoEvents calls.")]
        IFake DoEvents { get; }

        [DispId(7)]
        [Description("Configures VBA.Interaction.Shell calls.")]
        IFake Shell { get; }

        [DispId(8)]
        [Description("Configures VBA.Interaction.SendKeys calls.")]
        IStub SendKeys { get; }

        [DispId(9)]
        [Description("Configures VBA.FileSystem.Kill calls.")]
        IStub Kill { get; }

        [DispId(10)]
        [Description("Configures VBA.FileSystem.MkDir calls.")]
        IStub MkDir { get; }

        [DispId(11)]
        [Description("Configures VBA.FileSystem.RmDir calls.")]
        IStub RmDir { get; }

        [DispId(12)]
        [Description("Configures VBA.FileSystem.ChDir calls.")]
        IStub ChDir { get; }

        [DispId(13)]
        [Description("Configures VBA.FileSystem.ChDrive calls.")]
        IStub ChDrive { get; }

        //[DispId(17)]
        //IFake CurDir { get; }
    }
}
