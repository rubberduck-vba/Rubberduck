using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.UI;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchReplacePresenter : IPresenter
    {
        private readonly VBE _vbe;
        private readonly IRegexSearchReplaceDialog _view;
        private readonly IRubberduckParser _parser;

        public RegexSearchReplacePresenter(VBE vbe, IRubberduckParser parser, IRegexSearchReplaceDialog view)
        {
            _vbe = vbe;
            _view = view;
            _parser = parser;

            _view.FindButtonClicked += _view_FindButtonClicked;
            _view.ReplaceButtonClicked += _view_ReplaceButtonClicked;
            _view.ReplaceAllButtonClicked += _view_ReplaceAllButtonClicked;
            _view.CancelButtonClicked += _view_CancelButtonClicked;
        }

        public void Show()
        {
            _view.ShowDialog();
        }

        public event EventHandler<IEnumerable<RegexSearchResult>> FindButtonResults;
        protected virtual void OnFindButtonResults(IEnumerable<RegexSearchResult> results)
        {
            var handler = FindButtonResults;
            if (handler != null)
            {
                handler(this, results);
            }
        }

        private void _view_FindButtonClicked(object sender, EventArgs e)
        {
            var regexSearchReplace = new RegexSearchReplace(_vbe, _parser, new CodePaneWrapperFactory());
            OnFindButtonResults(regexSearchReplace.Search(_view.SearchPattern, _view.Scope));
        }

        private void _view_ReplaceButtonClicked(object sender, EventArgs e)
        {
            var regexSearchReplace = new RegexSearchReplace(_vbe, _parser, new CodePaneWrapperFactory());
            regexSearchReplace.Replace(_view.SearchPattern, _view.ReplacePattern, _view.Scope);
        }

        private void _view_ReplaceAllButtonClicked(object sender, EventArgs e)
        {
            var regexSearchReplace = new RegexSearchReplace(_vbe, _parser, new CodePaneWrapperFactory());
            regexSearchReplace.ReplaceAll(_view.SearchPattern, _view.ReplacePattern, _view.Scope);
        }

        private void _view_CancelButtonClicked(object sender, EventArgs e)
        {
            _view.Close();
        }
    }
}