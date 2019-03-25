using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.IdentifierReferences;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Resources;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.ParserErrors
{
    public interface IParserErrorsPresenter
    {
        void Show();
        void AddError(ParseErrorEventArgs error);
        void Clear();
    }

    public class ParserErrorsPresenter : DockableToolwindowPresenter, IParserErrorsPresenter
    {
        private readonly ISelectionService _selectionService;

        public ParserErrorsPresenter(IVBE vbe, IAddIn addin, ISelectionService selectionService) 
            : base(vbe, addin, new SimpleListControl(RubberduckUI.ParseErrors_Caption), null)
        {
            _selectionService = selectionService;
            _source = new BindingList<ParseErrorListItem>();
            var control = UserControl as SimpleListControl;
            Debug.Assert(control != null);
            control.Navigate += Control_Navigate;
        }

        private void Control_Navigate(object sender, ListItemActionEventArgs e)
        {
            var selection = (ParseErrorListItem) e.SelectedItem;
            selection.Navigate(_selectionService);
        }

        private readonly IBindingList _source;

        public void AddError(ParseErrorEventArgs error)
        {
            _source.Add(new ParseErrorListItem(error));
            var control = UserControl as SimpleListControl;
            Debug.Assert(control != null);

            if (control.InvokeRequired)
            {
                control.Invoke((MethodInvoker) delegate
                {
                    control.ResultBox.DataSource = _source;
                    control.ResultBox.DisplayMember = "Value";
                    control.Refresh();
                });
            }
            else
            {
                control.ResultBox.DataSource = _source;
                control.ResultBox.DisplayMember = "Value";
                control.Refresh();
            }
        }

        public void Clear()
        {
            _source.Clear();
        }
    }
}
