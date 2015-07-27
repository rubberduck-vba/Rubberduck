using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Navigations.RegexSearchReplace;

namespace Rubberduck.UI.RegexSearchReplace
{
    public partial class RegexSearchReplace : Form, IRegexSearchReplaceView
    {
        public string SearchPattern { get; private set; }
        public string ReplacePattern { get; private set; }
        public RegexSearchReplaceScope Scope { get; private set; }

        public RegexSearchReplace()
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
    }
}
