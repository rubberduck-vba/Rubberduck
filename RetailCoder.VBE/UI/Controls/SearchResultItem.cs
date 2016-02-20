using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Controls
{
    public class SearchResultItem : ViewModelBase, INavigateSource
    {
        private readonly NavigateCodeEventArgs _navigateArgs;
        private readonly QualifiedMemberName _member;
        private readonly Selection _selection;
        private string _resultText;

        public SearchResultItem(QualifiedMemberName member, Selection selection, string resultText)
        {
            _navigateArgs = new NavigateCodeEventArgs(member.QualifiedModuleName, selection);
            _member = member;
            _selection = selection;
            _resultText = resultText;
        }

        public QualifiedMemberName QualifiedMemberName { get { return _member; }}
        public Selection Selection { get { return _selection; } }

        public string ResultText
        {
            get { return _resultText; }
            set
            {
                if (_resultText != value)
                {
                    _resultText = value;
                    OnPropertyChanged();
                }
            }
        }
        
        public NavigateCodeEventArgs GetNavigationArgs()
        {
            return _navigateArgs;
        }
    }
}
