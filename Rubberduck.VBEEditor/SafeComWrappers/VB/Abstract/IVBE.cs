using Rubberduck.VBEditor.Host;
using System;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    // ReSharper disable once InconsistentNaming
    public interface IVBE : ISafeComWrapper, IEquatable<IVBE>
    {
        VBEKind Kind { get; }
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
        IReadOnlyDictionary<CommandBarSite, CommandBarLocation> CommandBarLocations { get; }
        QualifiedSelection? GetActiveSelection();

        bool IsInDesignMode { get; }
        int ProjectsCount { get; }
        ISourceCodeHandler SourceCodeHandler { get; }
    }
}