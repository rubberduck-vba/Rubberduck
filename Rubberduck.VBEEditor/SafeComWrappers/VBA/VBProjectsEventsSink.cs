using System;
using System.Diagnostics.CodeAnalysis;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class VBProjectsEventsSink : VB._dispVBProjectsEvents, IVBProjectsEventsSink
    {
        public event EventHandler<DispatcherEventArgs<IVBProject>> ProjectAdded;
        public void ItemAdded(VB.VBProject VBProject)
        {
            OnDispatch(ProjectAdded, VBProject);
        }

        public event EventHandler<DispatcherEventArgs<IVBProject>> ProjectRemoved;
        public void ItemRemoved(VB.VBProject VBProject)
        {
            OnDispatch(ProjectRemoved, VBProject);
        }

        public event EventHandler<DispatcherRenamedEventArgs<IVBProject>> ProjectRenamed;
        public void ItemRenamed(VB.VBProject VBProject, string OldName)
        {
            var handler = ProjectRenamed;
            if (handler != null && VBProject.Protection == VB.vbext_ProjectProtection.vbext_pp_none)
            {
                handler.Invoke(this, new DispatcherRenamedEventArgs<IVBProject>(new VBProject(VBProject), OldName));
            }
        }

        public event EventHandler<DispatcherEventArgs<IVBProject>> ProjectActivated;
        public void ItemActivated(VB.VBProject VBProject)
        {
            OnDispatch(ProjectActivated, VBProject);
        }

        private void OnDispatch(EventHandler<DispatcherEventArgs<IVBProject>> dispatched, VB.VBProject project)
        {
            var handler = dispatched;
            if (handler != null && project.Protection == VB.vbext_ProjectProtection.vbext_pp_none)
            {
                handler.Invoke(this, new DispatcherEventArgs<IVBProject>(new VBProject(project)));
            }
        }
    }
}
