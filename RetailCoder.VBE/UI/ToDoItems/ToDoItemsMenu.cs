using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.UI.ToDoItems
{
    [ComVisible(false)]
    public class ToDoItemsMenu 
    {
        private readonly VBE _vbe;
        private readonly AddIn _addIn;
        private readonly ToDoListSettings _settings;
        private readonly IRubberduckParser _parser;

        private CommandBarButton _todoItemsButton;

        public ToDoItemsMenu(VBE vbe, AddIn addInInstance, ToDoListSettings settings, IRubberduckParser parser)
        {
            _vbe = vbe;
            _addIn = addInInstance;
            _settings = settings;
            _parser = parser;
        }

        public void Initialize(CommandBarControls menuControls)
        {
            _todoItemsButton = menuControls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            Debug.Assert(_todoItemsButton != null);

            _todoItemsButton.Caption = "&ToDo Items";

            const int clipboardWithCheck = 837;
            _todoItemsButton.FaceId = clipboardWithCheck;
            _todoItemsButton.Click += OnShowTaskListButtonClick;
        }

        private void OnShowTaskListButtonClick(CommandBarButton ctrl, ref bool CancelDefault)
        {
            var markers = _settings.ToDoMarkers;
            var presenter = new ToDoExplorerDockablePresenter(_parser, markers, _vbe, _addIn);
            presenter.Show();
        }
    }
}
