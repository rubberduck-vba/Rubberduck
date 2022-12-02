using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

namespace Rubberduck.UnitTesting
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IFakesProviderGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public interface IFakesProvider
    {
        [DispId(1)]
        [Description("Configures VBA.Interaction.MsgBox calls.")]
        IFake MsgBox { get; }

        [DispId(2)]
        [Description("Configures VBA.Interaction.InputBox calls.")]
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

        [DispId(14)]
        [Description("Configures VBA.FileSystem.CurDir calls.")]
        IFake CurDir { get; }

        [DispId(15)]
        [Description("Configures VBA.DateTime.Now calls.")]
        IFake Now { get; }

        [DispId(16)]
        [Description("Configures VBA.DateTime.Time calls.")]
        IFake Time { get; }

        [DispId(17)]
        [Description("Configures VBA.DateTime.Date calls.")]
        IFake Date { get; }

        [DispId(18)]
        [Description("Configures VBA.Math.Rnd calls.")]
        IFake Rnd { get; }

        [DispId(19)]
        [Description("Configures VBA.Interaction.DeleteSetting calls.")]
        IStub DeleteSetting { get; }

        [DispId(20)]
        [Description("Configures VBA.Interaction.SaveSetting calls.")]
        IStub SaveSetting { get; }

        [DispId(21)]
        [Description("Configures VBA.Interaction.GetSetting calls.")]
        IFake GetSetting { get; }

        [DispId(22)]
        [Description("Configures VBA.Math.Randomize calls.")]
        IStub Randomize { get; }

        [DispId(23)]
        [Description("Configures VBA.Interaction.GetAllSettings calls.")]
        IFake GetAllSettings { get; }

        [DispId(24)]
        [Description("Configures VBA.FileSystem.SetAttr calls.")]
        IStub SetAttr { get; }

        [DispId(25)]
        [Description("Configures VBA.FileSystem.GetAttr calls.")]
        IFake GetAttr { get; }

        [DispId(26)]
        [Description("Configures VBA.FileSystem.FileLen calls.")]
        IFake FileLen { get; }

        [DispId(27)]
        [Description("Configures VBA.FileSystem.FileDateTime calls.")]
        IFake FileDateTime { get; }

        [DispId(28)]
        [Description("Configures VBA.FileSystem.FreeFile calls.")]
        IFake FreeFile { get; }

        [DispId(29)]
        [Description("Configures VBA.Information.IMEStatus calls.")]
        IFake IMEStatus { get; }

        [DispId(30)]
        [Description("Configures VBA.FileSystem.Dir calls.")]
        IFake Dir { get; }

        [DispId(31)]
        [Description("Configures VBA.FileSystem.FileCopy calls.")]
        IStub FileCopy { get; }


        [DispId(255)]
        [Description("Gets an interface exposing the parameter names for all parameterized fakes.")]
        IParams Params { get; }
    }
}
