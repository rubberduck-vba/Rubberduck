using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

// ReSharper disable InconsistentNaming
// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.

namespace Rubberduck.UnitTesting
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IParamsGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public interface IParams
    {
        [DispId(1)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.Interaction.MsgBox' function.")]
        IMsgBoxParams MsgBox { get; }

        [DispId(2)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.Interaction.InputBox' function.")]
        IInputBoxParams InputBox { get; }

        [DispId(3)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.Interaction.Environ' function.")]
        IEnvironParams Environ { get; }

        [DispId(4)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.Interaction.Shell' function.")]
        IShellParams Shell { get; }

        [DispId(5)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.Interaction.SendKeys' function.")]
        ISendKeysParams SendKeys { get; }

        [DispId(6)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.Kill' function.")]
        IKillParams Kill { get; }

        [DispId(7)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.FileCopy' statement.")]
        IFileCopyParams FileCopy { get; }

        [DispId(8)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.FreeFile' function.")]
        IFreeFileParams FreeFile { get; }

        [DispId(9)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.GetAttr' function.")]
        IGetAttrParams GetAttr { get; }

        [DispId(10)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.SetAttr' statement.")]
        ISetAttrParams SetAttr { get; }

        [DispId(11)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.FileLen' function.")]
        IFileLenParams FileLen { get; }

        [DispId(12)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.FileDateTime' function.")]
        IFileDateTimeParams FileDateTime { get; }

        [DispId(13)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.Dir' function.")]
        IDirParams Dir { get; }

        [DispId(14)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.CurDir' function.")]
        ICurDirParams CurDir { get; }

        [DispId(15)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.ChDir' statement.")]
        IChDirParams ChDir { get; }

        [DispId(16)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.ChDrive' statement.")]
        IChDriveParams ChDrive { get; }

        [DispId(17)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.MkDir' statement.")]
        IMkDirParams MkDir { get; }

        [DispId(18)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.FileSystem.RmDir' statement.")]
        IRmDirParams RmDir { get; }

        [DispId(19)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.Interaction.SaveSetting' statement.")]
        ISaveSettingParams SaveSetting { get; }

        [DispId(20)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.Interaction.DeleteSetting' statement.")]
        IDeleteSettingParams DeleteSetting { get; }

        [DispId(21)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.Math.Randomize' statement.")]
        IRandomizeParams Randomize { get; }

        [DispId(22)]
        [Description("Gets an interface exposing the parameter names for the 'VBA.Math.Rnd' function.")]
        IRndParams Rnd { get; }
    }
}
