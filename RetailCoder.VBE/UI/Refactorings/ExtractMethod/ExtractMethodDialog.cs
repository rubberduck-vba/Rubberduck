using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    public interface IDialogView
    {
        event EventHandler CancelButtonClicked;
        void OnCancelButtonClicked();

        event EventHandler OkButtonClicked;
        void OnOkButtonClicked();

        DialogResult ShowDialog();
    }

    public interface IExtractMethodDialog : IDialogView
    {
        event EventHandler<ValueChangedEventArgs<IdentifierNode>> MethodReturnValueChanged;
        void OnMethodReturnValueChanged(IdentifierNode newValue);

        event EventHandler<ValueChangedEventArgs<VBAccessibility>> MethodAccessibilityChanged;
        void OnMethodAccessibilityChanged(VBAccessibility newValue);

        event EventHandler<ValueChangedEventArgs<string>> MethodNameChanged;
        void OnMethodNameChanged(string newValue);

        string Preview { get; set; }
        IEnumerable<ExtractedParameter> Parameters { get; set; }

        string MethodName { get; set; }
    }

    public partial class ExtractMethodDialog : Form, IExtractMethodDialog
    {
        public ExtractMethodDialog()
        {
            InitializeComponent();
            RegisterViewEvents();

            MethodParametersGrid.DataSource = _parameters;
        }

        private void RegisterViewEvents()
        {
            OkButton.Click += OkButtonOnClick;
            CancelButton.Click += CancelButton_Click;

            MethodNameBox.TextChanged += MethodNameBox_TextChanged;
            MethodAccessibilityCombo.SelectedIndexChanged += MethodAccessibilityCombo_SelectedIndexChanged;
            MethodReturnValueCombo.SelectedIndexChanged += MethodReturnValueCombo_SelectedIndexChanged;
        }

        public event EventHandler<ValueChangedEventArgs<IdentifierNode>> MethodReturnValueChanged;

        public void OnMethodReturnValueChanged(IdentifierNode newValue)
        {
            var args = new ValueChangedEventArgs<IdentifierNode>(newValue);
            OnViewEvent(MethodReturnValueChanged, args);
        }

        private void MethodReturnValueCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            OnMethodReturnValueChanged((IdentifierNode)MethodReturnValueCombo.SelectedItem);
        }

        public event EventHandler<ValueChangedEventArgs<VBAccessibility>> MethodAccessibilityChanged;

        public void OnMethodAccessibilityChanged(VBAccessibility newValue)
        {
            var args = new ValueChangedEventArgs<VBAccessibility>(newValue);
            OnViewEvent(MethodAccessibilityChanged, args);
        }

        private void MethodAccessibilityCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            OnMethodAccessibilityChanged((VBAccessibility)MethodAccessibilityCombo.SelectedItem);
        }

        public event EventHandler<ValueChangedEventArgs<string>> MethodNameChanged;

        public void OnMethodNameChanged(string newValue)
        {
            var args = new ValueChangedEventArgs<string>(newValue);
            OnViewEvent(MethodNameChanged, args);
        }

        private void MethodNameBox_TextChanged(object sender, EventArgs e)
        {
            OnMethodNameChanged(MethodNameBox.Text);
        }

        private void OnViewEvent(EventHandler target, EventArgs args = null)
        {
            var handler = target;
            if (handler == null)
            {
                return;
            }

            handler(this, args ?? EventArgs.Empty);
        }

        private void OnViewEvent<T>(EventHandler<T> target, T args)
            where T : EventArgs
        {
            var handler = target;
            if (handler == null)
            {
                return;
            }

            handler(this, args);
        }

        public event EventHandler CancelButtonClicked;
        
        public void OnCancelButtonClicked()
        {
            OnViewEvent(CancelButtonClicked);
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            OnCancelButtonClicked();
        }

        public event EventHandler OkButtonClicked;

        public void OnOkButtonClicked()
        {
            OnViewEvent(OkButtonClicked);
        }

        private void OkButtonOnClick(object sender, EventArgs e)
        {
            OnOkButtonClicked();
        }

        private string _preview;
        public string Preview
        {
            get { return _preview; }
            set
            {
                _preview = value;
                PreviewBox.Text = _preview;
            }
        }

        private BindingList<ExtractedParameter> _parameters;
        public IEnumerable<ExtractedParameter> Parameters
        {
            get { return _parameters; }
            set
            {
                _parameters = new BindingList<ExtractedParameter>(value.ToList());
                MethodParametersGrid.DataSource = _parameters;
                MethodParametersGrid.Refresh();
            }
        }

        public string MethodName
        {
            get { return MethodNameBox.Text; }
            set { MethodNameBox.Text = value; }
        }
    }
}
