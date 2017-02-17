using System;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    // ReSharper disable once InconsistentNaming
    public interface IVBE : ISafeComWrapper, IEquatable<IVBE>
    {
        string Version { get; }
        object HardReference { get; }
        IWindow ActiveWindow { get; }
        ICodePane ActiveCodePane { get; set; }
        IVBProject ActiveVBProject { get; set; }
        IVBComponent SelectedVBComponent { get; }
        IWindow MainWindow { get; }
        IAddIns AddIns { get; }
        IVBProjects VBProjects { get; }
        ICodePanes CodePanes { get; }
        ICommandBars CommandBars { get; }
        IWindows Windows { get; }

        IHostApplication HostApplication();
        IWindow ActiveMDIChild();

        bool IsInDesignMode { get; }
    }
}