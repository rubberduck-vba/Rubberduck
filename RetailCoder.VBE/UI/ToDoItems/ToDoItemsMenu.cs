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
        private readonly IToDoExplorerWindow _userControl;
        private readonly ToDoExplorerDockablePresenter _presenter; //if presenter goes out of scope, so does it's toolwindow Issue #169

        private CommandBarButton _todoItemsButton;

        public ToDoItemsMenu(VBE vbe, AddIn addInInstance, ToDoListSettings settings, IRubberduckParser parser)
            :base(vbe, addInInstance)
        {
            _settings = settings;
            _parser = parser;
            //todo: inject dependencies
            _userControl = new ToDoExplorerWindow();
            _presenter = new ToDoExplorerDockablePresenter(_parser, _settings.ToDoMarkers, this.IDE, this.addInInstance, _userControl);
        }

        public void Initialize(CommandBarPopup menu)
        {
            const int clipboardWithCheck = 837;
            _todoItemsButton = AddButton(menu, "&ToDo Items", false, new CommandBarButtonClickEvent(OnShowTaskListButtonClick), clipboardWithCheck);
        }

        private void OnShowTaskListButtonClick(CommandBarButton ctrl, ref bool CancelDefault)
        {
            _presenter.Show();
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
