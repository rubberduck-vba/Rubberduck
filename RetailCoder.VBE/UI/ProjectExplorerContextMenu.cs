using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI
{
    public class ProjectExplorerContextMenu : Menu
    {
        private readonly VBE _vbe;
        private readonly IRubberduckParser _parser;

        // ReSharper disable once NotAccessedField.Local
        private CommandBarButton _findAllReferences;
        // ReSharper disable once NotAccessedField.Local
        private CommandBarButton _findAllImplementations;
        // ReSharper disable once NotAccessedField.Local
        private CommandBarButton _rename;
        // ReSharper disable once NotAccessedField.Local
        private CommandBarButton _inspect;
        // ReSharper disable once NotAccessedField.Local
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

            _findAllReferences = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem) as CommandBarButton;
            _findAllReferences.Caption = RubberduckUI.ProjectExplorerContextMenu_FindAllReferences;
            _findAllReferences.BeginGroup = true;
            _findAllReferences.Click += OnFindAllReferencesClick;

            _findAllImplementations = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem + 1) as CommandBarButton;
            _findAllImplementations.Caption = RubberduckUI.ProjectExplorerContextMenu_FindAllImplementations;
            _findAllImplementations.Click += OnFindAllImplementationsClick;

            _rename = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem + 2) as CommandBarButton;
            _rename.Caption = RubberduckUI.ProjectExplorerContextMenu_Rename;
            _rename.Click += OnRenameClick;

            _inspect = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem + 3) as CommandBarButton;
            _inspect.Caption = RubberduckUI.ProjectExplorerContextMenu_Inspect;
            _inspect.Click += OnInspectClick;

            _runAllTests = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem + 4) as CommandBarButton;
            _runAllTests.Caption = RubberduckUI.ProjectExplorerContextMenu_RunAllTests;
            _runAllTests.Click += OnRunAllTestsClick;
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnFindAllReferencesClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ContextMenuFindReferences(this, EventArgs.Empty);
        }

        public event EventHandler<NavigateCodeEventArgs> FindReferences;
        private void ContextMenuFindReferences(object sender, EventArgs e)
        {
            var progress = new ParsingProgressPresenter();
            var results = progress.Parse(_parser, _vbe.ActiveVBProject);

            var clsName = _vbe.SelectedVBComponent.Name;

            var clsDeclaration =
                results.Declarations.Items.FirstOrDefault(item => item.DeclarationType == DeclarationType.Class
                                                               && item.IdentifierName == clsName
                                                               && item.Project.Equals(_vbe.ActiveVBProject));

            var handler = FindReferences;
            if (handler != null)
            {
                handler(this, new NavigateCodeEventArgs(clsDeclaration));
            }
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnFindAllImplementationsClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ContextMenuFindImplementations(this, EventArgs.Empty);
        }

        public event EventHandler<NavigateCodeEventArgs> FindImplementations;
        private void ContextMenuFindImplementations(object sender, EventArgs e)
        {
            var progress = new ParsingProgressPresenter();
            var results = progress.Parse(_parser, _vbe.ActiveVBProject);

            var clsName = _vbe.SelectedVBComponent.Name;

            var clsDeclaration =
                results.Declarations.Items.FirstOrDefault(item => item.DeclarationType == DeclarationType.Class
                                                               && item.IdentifierName == clsName
                                                               && item.Project.Equals(_vbe.ActiveVBProject));

            var handler = FindImplementations;
            if (handler != null)
            {
                handler(this, new NavigateCodeEventArgs(clsDeclaration));
            }
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnRenameClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var progress = new ParsingProgressPresenter();
            var results = progress.Parse(_parser, _vbe.ActiveVBProject);

            var clsName = _vbe.SelectedVBComponent.Name;

            var clsDeclaration =
                results.Declarations.Items.FirstOrDefault(item => item.DeclarationType == DeclarationType.Class
                                                               && item.IdentifierName == clsName
                                                               && item.Project.Equals(_vbe.ActiveVBProject));

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(_vbe, view, results);
                var refactoring = new RenameRefactoring(factory);
                refactoring.Refactor(clsDeclaration);
            }
        }

        private void OnInspectClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ContextMenuRunInspections(this, EventArgs.Empty);
        }

        public event EventHandler RunInspections;
        private void ContextMenuRunInspections(object sender, EventArgs e)
        {
            var handler = RunInspections;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnRunAllTestsClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {

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
                _findAllReferences.Click -= OnFindAllReferencesClick;
                _findAllReferences.Delete();
            }

            if (_findAllImplementations != null)
            {
                _findAllImplementations.Click -= OnFindAllImplementationsClick;
                _findAllImplementations.Delete();
            }

            if (_rename != null)
            {
                _rename.Click -= OnRenameClick;
                _rename.Delete();
            }

            if (_inspect != null)
            {
                _inspect.Click -= OnInspectClick;
                _inspect.Delete();
            }

            if (_runAllTests != null)
            {
                _runAllTests.Click -= OnRunAllTestsClick;
                _runAllTests.Delete();
            }

            _disposed = true;
        }
    }
}
