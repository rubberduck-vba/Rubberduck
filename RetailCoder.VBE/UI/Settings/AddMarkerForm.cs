using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;

namespace Rubberduck.UI.Settings
{
    public partial class AddMarkerForm : Form, IAddTodoMarkerView
    {
        public AddMarkerForm()
        {
            InitializeComponent();

            IsValidMarker = false;

            OkButton.Text = RubberduckUI.OK_AllCaps;
            CancelButton.Text = RubberduckUI.CancelButtonText;
            TodoMarkerTextBoxLabel.Text = RubberduckUI.TodoSettings_TextLabel;
            TodoMarkerPriorityComboBoxLabel.Text = RubberduckUI.TodoSettings_PriorityLabel;

            OkButton.Click += OKButtonClick;
            CancelButton.Click += CancelButtonClick;
            TodoMarkerTextBox.TextChanged += MarkerTextChanged;
            TodoMarkerPriorityComboBox.DataSource = TodoLabels();
        }

        public List<ToDoMarker> TodoMarkers { get; set; }

        public string MarkerText
        {
            get { return TodoMarkerTextBox.Text; }
            set { TodoMarkerTextBox.Text = value; }
        }

        private bool _isValidMarker;
        public bool IsValidMarker
        {
            get { return _isValidMarker; }
            set
            {
                _isValidMarker = value;

                InvalidNameValidationIcon.Visible = !_isValidMarker;
                OkButton.Enabled = _isValidMarker;
            }
        }

        private List<string> TodoLabels()
        {
            return (from object priority in Enum.GetValues(typeof(TodoPriority))
                    select
                    RubberduckUI.ResourceManager.GetString("ToDoPriority_" + priority, RubberduckUI.Culture))
                    .ToList();
        }

        public TodoPriority MarkerPriority
        {
            get
            {
                return Enum.GetValues(typeof (TodoPriority))
                    .Cast<TodoPriority>()
                    .FirstOrDefault(
                        p =>
                            Equals(TodoMarkerPriorityComboBox.SelectedItem, RubberduckUI.ResourceManager.GetString("ToDoPriority_" + p, RubberduckUI.Culture)));
            }
            set
            {
                TodoMarkerPriorityComboBox.SelectedItem = RubberduckUI.ResourceManager.GetString("ToDoPriority_" + value, RubberduckUI.Culture);

            }
        }

        public event EventHandler AddMarker;
        private void OKButtonClick(object sender, EventArgs e)
        {
            var handler = AddMarker;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        public event EventHandler Cancel;
        private void CancelButtonClick(object sender, EventArgs e)
        {
            var handler = Cancel;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        public event EventHandler TextChanged;
        private void MarkerTextChanged(object sender, EventArgs e)
        {
            var handler = TextChanged;
            if (handler != null)
            {
                handler(sender, e);
            }
        }
    }
}
