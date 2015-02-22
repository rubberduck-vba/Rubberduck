using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.VBA;
using System;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;


namespace Rubberduck.UI.ToDoItems
{
    public class ToDoItemsMenu : Menu
    {
        private readonly ToDoListSettings _settings;
        private readonly IRubberduckParser _parser;
        private IToDoExplorerWindow _userControl;

        private CommandBarButton _todoItemsButton;

        public ToDoItemsMenu(VBE vbe, AddIn addInInstance, ToDoListSettings settings, IRubberduckParser parser)
            :base(vbe, addInInstance)
        {
            _settings = settings;
            _parser = parser;
        }

        public void Initialize(CommandBarPopup menu)
        {
            const int clipboardWithCheck = 837;
            _todoItemsButton = AddButton(menu, "&ToDo Items", false, new CommandBarButtonClickEvent(OnShowTaskListButtonClick), clipboardWithCheck);
        }

        private void OnShowTaskListButtonClick(CommandBarButton ctrl, ref bool CancelDefault)
        {
            var markers = _settings.ToDoMarkers;
            if (_userControl == null)
            {
                _userControl = new ToDoExplorerWindow();
            }
            var presenter = new ToDoExplorerDockablePresenter(_parser, markers, this.IDE, this.addInInstance, _userControl);
            presenter.Show();
        }

        bool disposed = false;
        protected override void Dispose(bool disposing)
        {
            if (disposed)
            {
                return;
            }

            if (disposing && _userControl != null)
            {
                var uc = (System.Windows.Forms.UserControl)_userControl;
                uc.Dispose();
            }

            disposed = true;

            base.Dispose();
        }
    }
}
