using System.Diagnostics.CodeAnalysis;
using System.Linq;

using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.VBIDEApi;
using Rubberduck.Parsing;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI
{
    internal class FormContextMenu
    {
        private readonly IRubberduckParser _parser;
        private readonly VBE _vbe;

        // ReSharper disable once NotAccessedField.Local
        private CommandBarButton _rename;

        public FormContextMenu(VBE vbe, IRubberduckParser parser)
        {
            _vbe = vbe;
            _parser = parser;
        }

        public void Initialize()
        {
            var beforeItem = _vbe.CommandBars["MSForms Control"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2558).Index;
            _rename = _vbe.CommandBars["MSForms Control"].Controls.Add(MsoControlType.msoControlButton, null, null, beforeItem, true) as CommandBarButton;
            _rename.BeginGroup = true;
            _rename.Caption = RubberduckUI.FormContextMenu_Rename;
            _rename.ClickEvent += OnRenameButtonClick;
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnRenameButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Rename();
        }

        private void Rename()
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, _vbe.ActiveVBProject);

            var designer = (dynamic) _vbe.SelectedVBComponent.Designer;

            foreach (var control in designer.Controls)
            {
                if (!control.InSelection) { continue; }

                var controlToRename =
                    result.Declarations.Items
                        .FirstOrDefault(item => item.IdentifierName == control.Name
                                                && item.ComponentName == _vbe.SelectedVBComponent.Name
                                                && _vbe.ActiveVBProject.Equals(item.Project));

                using (var view = new RenameDialog())
                {
                    var factory = new RenamePresenterFactory(_vbe, view, result);
                    var refactoring = new RenameRefactoring(factory);
                    refactoring.Refactor(controlToRename);
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) { return; }

            if (_rename != null)
            {
                _rename.ClickEvent -= OnRenameButtonClick;
                _rename.Delete();
            }
        }
    }
}
