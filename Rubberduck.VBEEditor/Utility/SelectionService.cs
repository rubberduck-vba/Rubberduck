using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Utility
{
    public class SelectionService : ISelectionService
    {
        private readonly IProjectsProvider _projectsProvider;
        private readonly IVBE _vbe;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public SelectionService(IVBE vbe, IProjectsProvider projectsProvider)
        {
            _vbe = vbe;
            _projectsProvider = projectsProvider;
        }

        public QualifiedModuleName ProjectExplorerSelection()
        {
            using (var component = _vbe.SelectedVBComponent)
            {
                return component.QualifiedModuleName;
            }
        }

        public QualifiedSelection? ActiveSelection()
        {
            using (var activeCodePane = _vbe.ActiveCodePane)
            {
                return activeCodePane?.GetQualifiedSelection();
            }
        }

        public ICollection<QualifiedModuleName> OpenModules()
        {
            var openModules = new HashSet<QualifiedModuleName>();
            using (var openCodePanes = _vbe.CodePanes)
            {
                if (openCodePanes == null || openCodePanes.IsWrappingNullReference)
                {
                    return new HashSet<QualifiedModuleName>();
                }

                foreach (var openCodePane in openCodePanes)
                {
                    using (openCodePane)
                    {

                        try
                        {
                            var qmn = openCodePane.QualifiedModuleName;
                            openModules.Add(qmn);
                        }
                        catch (COMException ex)
                        {
                            //For some reason, we sometimes get a COM Exception for 'invalid callee' here. So, we swallow it to avoid rendering rewrites unusable in that case.
                            //See issue #5242 at https://github.com/rubberduck-vba/Rubberduck/issues/5242
                            //TODO: Find the root cause of the sporadic exception and deal with it.
                            _logger.Warn(ex, "Encountered an exception while getting the qualified module name of all open code panes. The current code pane will be skipped.");
                        }
                    }
                }
            }

            return openModules;
        }

        public Selection? Selection(QualifiedModuleName module)
        {
            var component = _projectsProvider.Component(module);
            if (component == null)
            {
                return null;
            }

            using (var codeModule = component.CodeModule)
            using (var codePane = codeModule.CodePane)
            {
                return codePane.Selection;
            }
        }

        public bool TryActivate(QualifiedModuleName module)
        {
            try
            {
                var component = _projectsProvider.Component(module);
                if (component == null)
                {
                    return false;
                }

                using (var codeModule = component.CodeModule)
                using(var codePane = codeModule.CodePane)
                {
                    _vbe.ActiveCodePane = codePane;
                }

                return true;
            }
            catch (Exception exception)
            {
                _logger.Debug(exception, $"Failed to activate the code pane of module {module}.");
                return false;
            }
        }

        public bool TrySetActiveSelection(QualifiedModuleName module, Selection selection)
        {
            var activeCodePane = _vbe.ActiveCodePane;

            if (!TryActivate(module))
            {
                return false;
            }

            if (!TrySetSelection(module, selection))
            {
                TryActivate(activeCodePane.QualifiedModuleName);
                return false;
            }

            return true;
        }

        public bool TrySetActiveSelection(QualifiedSelection selection)
        {
            return TrySetActiveSelection(selection.QualifiedName, selection.Selection);
        }

        public bool TrySetSelection(QualifiedModuleName module, Selection selection)
        {
            try
            {
                var component = _projectsProvider.Component(module);
                if (component == null)
                {
                    return false;
                }

                using (var codeModule = component.CodeModule)
                using (var codePane = codeModule.CodePane)
                {
                    codePane.Selection = selection;
                }

                return true;
            }
            catch (Exception exception)
            {
                _logger.Debug(exception, $"Failed to set the selection of module {module} to {selection}.");
                return false;
            }
        }

        public bool TrySetSelection(QualifiedSelection selection)
        {
            return TrySetSelection(selection.QualifiedName, selection.Selection);
        }
    }
}