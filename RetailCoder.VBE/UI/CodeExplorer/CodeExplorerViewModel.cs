using System;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerViewModel : ViewModelBase
    {
        private readonly RubberduckParserState _state;
        private readonly ObservableCollection<ExplorerItemViewModel> _children;

        public CodeExplorerViewModel(RubberduckParserState state, ObservableCollection<ExplorerItemViewModel> children)
        {
            _state = state;
            _children = children;
        }

        private bool _isBusy;
        public bool IsBusy { get { return _isBusy; } set { _isBusy = value; OnPropertyChanged(); } }


    }

    public class ExplorerItemViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;
        private readonly ConcurrentStack<ExplorerItemViewModel> _children = new ConcurrentStack<ExplorerItemViewModel>(); 

        public ExplorerItemViewModel(Declaration declaration)
        {
            _declaration = declaration;
        }

        public void AddChild(ExplorerItemViewModel declaration)
        {
            _children.Push(declaration);
        }

        public void Clear()
        {
            _children.Clear();
        }

        public Declaration Declaration { get { return _declaration; } }
    }
}
