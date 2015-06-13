﻿using NetOffice.OfficeApi;
using NetOffice.VBIDEApi;
using CommandBarButtonClickEvent = NetOffice.OfficeApi.CommandBarButton_ClickEventHandler;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsMenu : Menu
    {
        private CommandBarButton _codeInspectionsButton;
        private readonly CodeInspectionsWindow _window;
        private readonly CodeInspectionsDockablePresenter _presenter; //if presenter goes out of scope, so does it's toolwindow Issue #169

        public CodeInspectionsMenu(VBE vbe, AddIn addIn, CodeInspectionsWindow view, CodeInspectionsDockablePresenter presenter)
            :base(vbe, addIn)
        {
            _window = view;
            _presenter = presenter;
        }

        public void Initialize(CommandBarPopup parentMenu)
        {
            _codeInspectionsButton = AddButton(parentMenu, RubberduckUI.RubberduckMenu_CodeInspections, false, new CommandBarButtonClickEvent(OnCodeInspectionsButtonClick));
        }

        public void Inspect()
        {
            _presenter.Show();
        }

        private void OnCodeInspectionsButtonClick(CommandBarButton ctrl, ref bool canceldefault)
        {
            Inspect();
        }

        bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (_disposed || !disposing)
            {
                return;
            }

            if (_window != null)
            {
                _window.Dispose();
            }

            _codeInspectionsButton.ClickEvent -= OnCodeInspectionsButtonClick;

            _disposed = true;
            base.Dispose(true);
        }
    }
}
