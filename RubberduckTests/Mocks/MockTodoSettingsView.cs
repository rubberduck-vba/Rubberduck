using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Mocks
{
    class MockTodoSettingsView : Rubberduck.UI.Settings.ITodoSettingsView
    {
        public Rubberduck.Config.TodoPriority ActiveMarkerPriority
        {
            get;
            set;
        }

        public string ActiveMarkerText
        {
            get;
            set;
        }

        public System.ComponentModel.BindingList<Rubberduck.Config.ToDoMarker> TodoMarkers
        {
            get;
            set;
        }

        private int selectedIndex;
        public int SelectedIndex
        {
            get { return selectedIndex; }
            set
            {
                selectedIndex = value;
                OnSelectionChanged(EventArgs.Empty);
            }
        }

        public event EventHandler RemoveMarker;

        public event EventHandler AddMarker;

        public event EventHandler SaveMarker;

        public event EventHandler SelectionChanged;
        protected virtual void OnSelectionChanged(EventArgs e)
        {
            EventHandler handler = SelectionChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
