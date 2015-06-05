using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.UI
{
    public class ProjectExplorerContextMenu
    {
        private readonly VBE _vbe;
        private readonly IRubberduckParser _parser;

        // ReSharper disable once NotAccessedField.Local
        private CommandBarButton _navigate;
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
        
        public ProjectExplorerContextMenu(VBE vbe, IRubberduckParser parser)
        {
            _vbe = vbe;
            _parser = parser;
        }

        public void Initialize()
        {
            var beforeItem = _vbe.CommandBars["Project Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2578).Index;
            _navigate = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem) as CommandBarButton;
            _navigate.BeginGroup = true;
            _navigate.Caption = RubberduckUI.ProjectExplorerContextMenu_Navigate;
            _navigate.Click += OnNavigateButtonClick;

            _findAllReferences = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem + 1) as CommandBarButton;
            _findAllReferences.Caption = RubberduckUI.ProjectExplorerContextMenu_FindAllReferences;
            _findAllReferences.Click += OnFindAllReferencesClick;

            _findAllImplementations = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem + 2) as CommandBarButton;
            _findAllImplementations.Caption = RubberduckUI.ProjectExplorerContextMenu_FindAllImplementations;
            _findAllImplementations.Click += OnFindAllImplementationsClick;

            _rename = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem + 3) as CommandBarButton;
            _rename.Caption = RubberduckUI.ProjectExplorerContextMenu_Rename;
            _rename.Click += OnRenameClick;

            _inspect = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem + 4) as CommandBarButton;
            _inspect.Caption = RubberduckUI.ProjectExplorerContextMenu_Inspect;
            _inspect.Click += OnInspectClick;

            _runAllTests = _vbe.CommandBars["Project Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem + 5) as CommandBarButton;
            _runAllTests.Caption = RubberduckUI.ProjectExplorerContextMenu_RunAllTests;
            _runAllTests.Click += OnRunAllTestsClick;
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnNavigateButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnFindAllReferencesClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnFindAllImplementationsClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnRenameClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnInspectClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnRunAllTestsClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {

        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing)
            {
                return;
            }

            if (_navigate != null)
            {
                _navigate.Click -= OnNavigateButtonClick;
                _navigate.Delete();
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
        }
    }
}
