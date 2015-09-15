using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Navigation.RegexSearchReplace;

namespace Rubberduck.UI.RegexSearchReplace
{
    public partial class RegexSearchReplaceDialog : Form, IRegexSearchReplaceDialog
    {
        public string SearchPattern { get { return SearchBox.Text.Replace(@"\\", @"\"); } }
        public string ReplacePattern { get { return ReplaceBox.Text; } }
        public RegexSearchReplaceScope Scope { get { return ConvertScopeLabelsToEnum(); } }

        public RegexSearchReplaceDialog()
        {
            InitializeComponent();

            InitializeCaptions();
            ScopeComboBox.DataSource = ScopeLabels();
        }
        
        private void InitializeCaptions()
        {
            Text = RubberduckUI.RegexSearchReplace_Caption;
            SearchLabel.Text = RubberduckUI.RegexSearchReplace_SearchLabel;
            ReplaceLabel.Text = RubberduckUI.RegexSearchReplace_ReplaceLabel;
            ScopeLabel.Text = RubberduckUI.RegexSearchReplace_ScopeLabel;

            FindButton.Text = RubberduckUI.RegexSearchReplace_FindButtonLabel;
            ReplaceButton.Text = RubberduckUI.RegexSearchReplace_ReplaceButtonLabel;
            ReplaceAllButton.Text = RubberduckUI.RegexSearchReplace_ReplaceAllButtonLabel;
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;
        }

        private List<string> ScopeLabels()
        {
            return (from object scope in Enum.GetValues(typeof(RegexSearchReplaceScope))
                    select
                    RubberduckUI.ResourceManager.GetString("RegexSearchReplaceScope_" + scope, RubberduckUI.Culture))
                    .ToList();
        }

        private RegexSearchReplaceScope ConvertScopeLabelsToEnum()
        {
            var scopes = from RegexSearchReplaceScope scope in Enum.GetValues(typeof(RegexSearchReplaceScope))
                         where ReferenceEquals(RubberduckUI.ResourceManager.GetString("RegexSearchReplaceScope_" + scope, RubberduckUI.Culture), ScopeComboBox.SelectedValue)
                         select scope;


            return scopes.First();
        }

        public event EventHandler<EventArgs> FindButtonClicked;
        protected virtual void OnFindButtonClicked(object sender, EventArgs e)
        {
            var handler = FindButtonClicked;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler<EventArgs> ReplaceButtonClicked;
        protected virtual void OnReplaceButtonClicked(object sender, EventArgs e)
        {
            var handler = ReplaceButtonClicked;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler<EventArgs> ReplaceAllButtonClicked;
        protected virtual void OnReplaceAllButtonClicked(object sender, EventArgs e)
        {
            var handler = ReplaceAllButtonClicked;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public event EventHandler<EventArgs> CancelButtonClicked;
        protected virtual void OnCancelButtonClicked(object sender, EventArgs e)
        {
            var handler = CancelButtonClicked;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }
    }
}
