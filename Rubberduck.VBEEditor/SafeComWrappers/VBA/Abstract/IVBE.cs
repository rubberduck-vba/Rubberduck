using System;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract
{
    // ReSharper disable once InconsistentNaming
    public interface IVBE : ISafeComWrapper, IEquatable<IVBE>
    {
        string Version { get; }
        IWindow ActiveWindow { get; }
        ICodePane ActiveCodePane { get; set; }
        VBProject ActiveVBProject { get; set; }
        VBComponent SelectedVBComponent { get; }
        IWindow MainWindow { get; }
        IAddIns AddIns { get; }
        VBProjects VBProjects { get; }
        ICodePanes CodePanes { get; }
        ICommandBars CommandBars { get; }
        IWindows Windows { get; }

        bool IsInDesignMode { get; }
    }
}