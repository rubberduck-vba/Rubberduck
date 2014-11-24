using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.VBA.Parser;

namespace Rubberduck.UI.ToDoItems
{
    internal class ToDoItemsMenu 
    {
        private readonly VBE _vbe;
        private readonly AddIn _addIn;
        private readonly ToDoListSettings _settings;
        private readonly Parser _parser;

        private CommandBarButton _todoItemsButton;
        public CommandBarButton ToDoItemsButton { get { return _todoItemsButton; } }

        public ToDoItemsMenu(VBE vbe, AddIn addInInstance, ToDoListSettings settings, Parser parser)
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
            _todoItemsButton.BeginGroup = true;

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
