using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.VBIDEApi;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI
{
    public class ProjectExplorerContextMenu : Menu
    {
        private readonly VBE _vbe;
        private readonly IRubberduckParser _parser;

        private CommandBarButton _findAllReferences;
        private CommandBarButton _findAllImplementations;
        private CommandBarButton _rename;
        private CommandBarButton _inspect;
        private CommandBarButton _runAllTests;
        
        public ProjectExplorerContextMenu(VBE vbe, AddIn addIn, IRubberduckParser parser)
            : base(vbe, addIn)
        {
            _vbe = vbe;
            _parser = parser;
        }

        public void Initialize()
        {
            var beforeItem = _vbe.CommandBars["Project Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2578).Index;

            _findAllReferences = (CommandBarButton)_vbe.CommandBars["Project Window"].Controls.Add(MsoControlType.msoControlButton, null, null, beforeItem, true);
            _findAllReferences.Caption = RubberduckUI.CodeExplorer_FindAllReferencesText;
            _findAllReferences.BeginGroup = true;
            _findAllReferences.ClickEvent += FindAllReferences_Click;

            _findAllImplementations = (CommandBarButton)_vbe.CommandBars["Project Window"].Controls.Add(MsoControlType.msoControlButton, null, null, beforeItem + 1, true);
            _findAllImplementations.Caption = RubberduckUI.CodeExplorer_FindAllImplementationsText;
            _findAllImplementations.ClickEvent += FindAllImplementations_Click;

            _rename = (CommandBarButton)_vbe.CommandBars["Project Window"].Controls.Add(MsoControlType.msoControlButton, null, null, beforeItem + 2, true);
            _rename.Caption = RubberduckUI.RefactorMenu_Rename;
            _rename.ClickEvent += Rename_Click;

            _inspect = (CommandBarButton)_vbe.CommandBars["Project Window"].Controls.Add(MsoControlType.msoControlButton, null, null, beforeItem + 3, true);
            _inspect.Caption = RubberduckUI.Inspect;
            _inspect.ClickEvent += Inspect_Click;

            _runAllTests = (CommandBarButton)_vbe.CommandBars["Project Window"].Controls.Add(MsoControlType.msoControlButton, null, null, beforeItem + 4, true);
            _runAllTests.Caption = RubberduckUI.CodeExplorer_RunAllTestsText;
            _runAllTests.ClickEvent += RunAllTests_Click;
        }

        private Declaration FindSelectedDeclaration()
        {
            VBProjectParseResult result;
            return FindSelectedDeclaration(out result);
        }

        private Declaration FindSelectedDeclaration(out VBProjectParseResult results)
        {
            var project = _vbe.ActiveVBProject;
            if (project == null)
            {
                results = null;
                return null;
            }

            var progress = new ParsingProgressPresenter();
            results = progress.Parse(_parser, _vbe.ActiveVBProject);

            var selection = _vbe.SelectedVBComponent;
            if (selection != null)
            {
                var componentName = selection.Name;
                var matches = results.Declarations[componentName].ToList();
                if (matches.Count == 1)
                {
                    return matches.Single();
                }

                var result = matches.SingleOrDefault(item =>
                    (item.DeclarationType == DeclarationType.Class || item.DeclarationType == DeclarationType.Module)
                    && item.Project == project);

                return result;
            }

            return results.Declarations[project.Name].SingleOrDefault(item =>
                item.DeclarationType == DeclarationType.Project && item.Project == project);
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void FindAllReferences_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var declaration = FindSelectedDeclaration();
            if (declaration == null)
            {
                return;
            }
            OnFindReferences(this, new NavigateCodeEventArgs(declaration));
        }

        public event EventHandler<NavigateCodeEventArgs> FindReferences;
        private void OnFindReferences(object sender, NavigateCodeEventArgs e)
        {
            var handler = FindReferences;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void FindAllImplementations_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var declaration = FindSelectedDeclaration();
            if (declaration == null)
            {
                return;
            }
            OnFindImplementations(this, new NavigateCodeEventArgs(declaration));
        }

        public event EventHandler<NavigateCodeEventArgs> FindImplementations;
        private void OnFindImplementations(object sender, NavigateCodeEventArgs e)
        {
            var handler = FindImplementations;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void Rename_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            VBProjectParseResult results;
            var declaration = FindSelectedDeclaration(out results);
            if (declaration == null)
            {
                return;
            }

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(_vbe, view, results);
                var refactoring = new RenameRefactoring(factory);
                refactoring.Refactor(declaration);
            }
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void Inspect_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnRunInspections(this, EventArgs.Empty);
        }

        public event EventHandler RunInspections;
        private void OnRunInspections(object sender, EventArgs e)
        {
            var handler = RunInspections;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void RunAllTests_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnRunAllTests(this, EventArgs.Empty);
        }

        public event EventHandler RunAllTests;
        private void OnRunAllTests(object sender, EventArgs e)
        {
            var handler = RunAllTests;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }
            
            if (_findAllReferences != null)
            {
                _findAllReferences.ClickEvent -= FindAllReferences_Click;
                _findAllReferences.Delete();
            }

            if (_findAllImplementations != null)
            {
                _findAllImplementations.ClickEvent -= FindAllImplementations_Click;
                _findAllImplementations.Delete();
            }

            if (_rename != null)
            {
                _rename.ClickEvent -= Rename_Click;
                _rename.Delete();
            }

            if (_inspect != null)
            {
                _inspect.ClickEvent -= Inspect_Click;
                _inspect.Delete();
            }

            if (_runAllTests != null)
            {
                _runAllTests.ClickEvent -= RunAllTests_Click;
                _runAllTests.Delete();
            }

            _disposed = true;
        }
    }
}
