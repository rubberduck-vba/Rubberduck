using Rubberduck.UI;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchReplaceModel : ViewModelBase
    {
        private string _searchPattern;
        public string SearchPattern { get { return _searchPattern; } set { _searchPattern = value; OnPropertyChanged(); } }

        private string _replacePattern;
        public string ReplacePattern { get { return _replacePattern; } set { _replacePattern = value; OnPropertyChanged(); } }

        private RegexSearchReplaceScope _searchScope;
        public RegexSearchReplaceScope SearchScope { get { return _searchScope; } set { _searchScope = value; OnPropertyChanged(); } }
    }
}
