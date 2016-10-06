using System;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    // ReSharper disable once InconsistentNaming
    public interface IVBE : ISafeComWrapper, IEquatable<IVBE>
    {
        string Version { get; }
        Window ActiveWindow { get; }
        CodePane ActiveCodePane { get; set; }
        VBProject ActiveVBProject { get; set; }
        VBComponent SelectedVBComponent { get; }
        Window MainWindow { get; }
        IAddIns AddIns { get; }
        VBProjects VBProjects { get; }
        CodePanes CodePanes { get; }
        ICommandBars CommandBars { get; }
        IWindows Windows { get; }

        bool IsInDesignMode { get; }
    }
}