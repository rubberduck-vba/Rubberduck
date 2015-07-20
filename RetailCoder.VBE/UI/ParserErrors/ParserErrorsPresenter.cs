using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.IdentifierReferences;

namespace Rubberduck.UI.ParserErrors
{
    public class ParseErrorListItem
    {
        private readonly ParseErrorEventArgs _error;

        public ParseErrorListItem(ParseErrorEventArgs error)
        {
            _error = error;
        }

        public string ProjectName { get { return _error.ProjectName; } }
        public string ComponentName { get { return _error.ComponentName; } }
        public int ErrorLine { get { return _error.Exception.LineNumber; } }
        public int ErrorColumn { get { return _error.Exception.Position; } }
        public string ErrorToken { get { return _error.Exception.OffendingSymbol.Text; } }
        public string Message { get { return _error.Exception.Message; } }

        public string Value { get { return ToString(); } }

        public void Navigate()
        {
            _error.Navigate();
        }

        public override string ToString()
        {
            return string.Format("{0}.{1} ({2},{3}): {4}", ProjectName, ComponentName, ErrorLine, ErrorColumn, Message);
        }
    }

    public interface IParserErrorsPresenter
    {
        void Show();
        void Clear();
        void AddError(ParseErrorEventArgs error);
    }

    public class ParserErrorsPresenter : DockablePresenterBase, IParserErrorsPresenter
    {
        public ParserErrorsPresenter(VBE vbe, AddIn addin) 
            : base(vbe, addin, new SimpleListControl(RubberduckUI.ParseErrors_Caption))
        {
            _source = new BindingList<ParseErrorListItem>();
            Control.Navigate += Control_Navigate;
        }

        void Control_Navigate(object sender, ListItemActionEventArgs e)
        {
            var selection = (ParseErrorListItem) e.SelectedItem;
            selection.Navigate();
        }

        private SimpleListControl Control { get { return (SimpleListControl) UserControl; } }

        private readonly IBindingList _source;

        public void AddError(ParseErrorEventArgs error)
        {
            _source.Add(new ParseErrorListItem(error));
            var control = Control;
            if (control.InvokeRequired)
            {
                control.Invoke((MethodInvoker) delegate
                {
                    Control.ResultBox.DataSource = _source;
                    Control.ResultBox.DisplayMember = "Value";
                    control.Refresh();
                });
            }
            else
            {
                Control.ResultBox.DataSource = _source;
                Control.ResultBox.DisplayMember = "Value";
                control.Refresh();
            }
        }

        public void Clear()
        {
            _source.Clear();
        }
    }
}
