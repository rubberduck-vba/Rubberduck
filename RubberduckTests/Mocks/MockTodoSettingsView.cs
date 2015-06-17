using System;
using System.Collections.Generic;
using System.ComponentModel;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Mocks
{
    class MockTodoSettingsView : ITodoSettingsView
    {
        public bool SaveEnabled { get; set; }
        public bool AddEnabled { get; set; }

        private TodoPriority activeMarkerPriority;
        public TodoPriority ActiveMarkerPriority
        {
            get { return activeMarkerPriority; }
            set
            {
                activeMarkerPriority = value;
                OnPriorityChanged(EventArgs.Empty);
            }
        }
        private string activeMarkerText;
        public string ActiveMarkerText
        {
            get { return activeMarkerText; }
            set
            {
                activeMarkerText = value;
                OnTextChanged(EventArgs.Empty);
            }
        }

        public BindingList<ToDoMarker> TodoMarkers { get; set; }

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

        public MockTodoSettingsView(List<ToDoMarker> markers)
        {
            this.TodoMarkers = new BindingList<ToDoMarker>(markers);
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

        public event EventHandler TextChanged;
        protected virtual void OnTextChanged(EventArgs e)
        {
            EventHandler handler = TextChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler PriorityChanged;
        protected virtual void OnPriorityChanged(EventArgs e)
        {
            EventHandler handler = TextChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
