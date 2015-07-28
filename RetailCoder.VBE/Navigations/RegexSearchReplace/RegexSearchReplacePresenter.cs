using System;
using System.Collections.Generic;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public class RegexSearchReplacePresenter
    {
        private readonly IRegexSearchReplaceView _view;
        private readonly RegexSearchReplaceModel _model;

        public RegexSearchReplacePresenter(IRegexSearchReplaceView view, RegexSearchReplaceModel model)
        {
            _view = view;
            _model = model;

            _view.FindButtonClicked += _view_FindButtonClicked;
            _view.ReplaceButtonClicked += _view_ReplaceButtonClicked;
            _view.ReplaceAllButtonClicked += _view_ReplaceAllButtonClicked;
            _view.CancelButtonClicked += _view_CancelButtonClicked;
        }

        public void Show()
        {
            _view.ShowDialog();
        }

        public event EventHandler<List<RegexSearchResult>> FindButtonResults;
        protected virtual void OnFindButtonResults(List<RegexSearchResult> results)
        {
            var handler = FindButtonResults;
            if (handler != null)
            {
                handler(this, results);
            }
        }

        private void _view_FindButtonClicked(object sender, EventArgs e)
        {
            var regexSearchReplace = new RegexSearchReplace(_model);
            OnFindButtonResults(regexSearchReplace.Find(_view.SearchPattern, _view.Scope));
        }

        private void _view_ReplaceButtonClicked(object sender, EventArgs e)
        {
            var regexSearchReplace = new RegexSearchReplace(_model);
            regexSearchReplace.Replace(_view.SearchPattern, _view.ReplacePattern, _view.Scope);
        }

        private void _view_ReplaceAllButtonClicked(object sender, EventArgs e)
        {
            var regexSearchReplace = new RegexSearchReplace(_model);
            regexSearchReplace.ReplaceAll(_view.SearchPattern, _view.ReplacePattern, _view.Scope);
        }

        private void _view_CancelButtonClicked(object sender, EventArgs e)
        {
            _view.Close();
        }
    }
}