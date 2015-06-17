using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public partial class AddMarkerForm : Form, IAddTodoMarkerView
    {
        public AddMarkerForm()
        {
            InitializeComponent();

            OkButton.Click += OKButtonClick;
            CancelButton.Click += CancelButtonClick;
            TodoMarkerTextBox.TextChanged += MarkerTextChanged;
        }

        public List<ToDoMarker> TodoMarkers { get; set; }

        public string MarkerText
        {
            get { return TodoMarkerTextBox.Text; }
            set { TodoMarkerTextBox.Text = value; }
        }

        private bool _isValidMarker = false;
        public bool IsValidMarker
        {
            get { return _isValidMarker; }
            set
            {
                _isValidMarker = value;

                InvalidNameValidationIcon.Visible = !_isValidMarker;
                OkButton.Enabled = !_isValidMarker;
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
