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

        public CodeExplorerViewModel(RubberduckParserState state)
        {
            _state = state;
            _state.StateChanged += ParserState_StateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;
        }

        private bool _isBusy;
        public bool IsBusy { get { return _isBusy; } set { _isBusy = value; OnPropertyChanged(); } }

        private void ParserState_StateChanged(object sender, EventArgs e)
        {
            IsBusy = _state.Status == ParserState.Parsing;
        }

        private void ParserState_ModuleStateChanged(object sender, Parsing.ParseProgressEventArgs e)
        {
            
        }
    }

    public class ExplorerProjectItemViewModel : ViewModelBase
    {
        
    }

    public class ExplorerComponentItemViewModel : ViewModelBase
    {
        
    }

    public class ExplorerMemberViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;
        private readonly ConcurrentStack<ExplorerMemberViewModel> _children = new ConcurrentStack<ExplorerMemberViewModel>(); 

        public ExplorerMemberViewModel(Declaration declaration)
        {
            _declaration = declaration;
        }

        public void AddChild(ExplorerMemberViewModel declaration)
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
