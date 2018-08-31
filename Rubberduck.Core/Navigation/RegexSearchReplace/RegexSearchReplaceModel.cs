using Rubberduck.UI;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchReplaceModel : ViewModelBase
    {
        private string _searchPattern;
        public string SearchPattern
        {
            get => _searchPattern;
            set
            {
                _searchPattern = value;
                OnPropertyChanged();
            } 
        }

        private string _replacePattern;
        public string ReplacePattern
        {
            get => _replacePattern;
            set
            {
                _replacePattern = value;
                OnPropertyChanged();
            } 
        }

        private RegexSearchReplaceScope _searchScope;
        public RegexSearchReplaceScope SearchScope
        {
            get => _searchScope;
            set
            {
                _searchScope = value;
                OnPropertyChanged(); 
                
            } 
        }
    }
}
