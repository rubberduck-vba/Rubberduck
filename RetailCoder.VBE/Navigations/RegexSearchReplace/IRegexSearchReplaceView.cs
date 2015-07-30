using System;
using System.Windows.Forms;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public interface IRegexSearchReplaceView
    {
        string SearchPattern { get; }
        string ReplacePattern { get; }
        RegexSearchReplaceScope Scope { get; }

        event EventHandler<EventArgs> FindButtonClicked;
        event EventHandler<EventArgs> ReplaceButtonClicked;
        event EventHandler<EventArgs> ReplaceAllButtonClicked;
        event EventHandler<EventArgs> CancelButtonClicked;

        DialogResult ShowDialog();
        void Close();
    }
}