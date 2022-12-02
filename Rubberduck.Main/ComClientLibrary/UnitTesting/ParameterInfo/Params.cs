using Rubberduck.Resources.Registration;
using Rubberduck.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UnitTesting
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.ParamsClassGuid),
        ProgId(RubberduckProgId.ParamsClassProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IParams)),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public class Params : IParams
    {
        public IMsgBoxParams MsgBox { get; } = new MsgBoxParams();
        public IInputBoxParams InputBox { get; } = new InputBoxParams();
        public IEnvironParams Environ { get; } = new EnvironParams();
        public IShellParams Shell { get; } = new ShellParams();
        public ISendKeysParams SendKeys { get; } = new SendKeysParams();
        public IKillParams Kill { get; } = new KillParams();
        public IFileCopyParams FileCopy { get; } = new FileCopyParams();
        public IFreeFileParams FreeFile { get; } = new FreeFileParams();
        public IGetAttrParams GetAttr { get; } = new GetAttrParams();
        public ISetAttrParams SetAttr { get; } = new SetAttrParams();
        public IFileLenParams FileLen { get; } = new FileLenParams();
        public IFileDateTimeParams FileDateTime { get; } = new FileDateTimeParams();
        public IDirParams Dir { get; } = new DirParams();
        public ICurDirParams CurDir { get; } = new CurDirParams();
        public IChDirParams ChDir { get; } = new ChDirParams();
        public IChDriveParams ChDrive { get; } = new ChDriveParams();
        public IMkDirParams MkDir { get; } = new MkDirParams();
        public IRmDirParams RmDir { get; } = new RmDirParams();
        public IDeleteSettingParams DeleteSetting { get; } = new DeleteSettingParams();
        public ISaveSettingParams SaveSetting { get; } = new SaveSettingParams();
        public IRandomizeParams Randomize { get; } = new RandomizeParams();
        public IRndParams Rnd { get; } = new RndParams();
    }
}
