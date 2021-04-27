using System.Linq;
using System.Windows.Input;
using Antlr4.Runtime;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Controls
{
    public class PeekDefinitionViewModel : ViewModelBase
    {
        internal PeekDefinitionViewModel()
        {
            // default constructor for xaml designer
        }

        public PeekDefinitionViewModel(ICodeExplorerNode node,
            ICommand findReferencesCommand,
            ICommand navigateCommand,
            ICommand closeCommand)
        :this(node.Declaration, findReferencesCommand, navigateCommand, closeCommand)
        {
            NavigateCommandParameter = node;
        }

        public PeekDefinitionViewModel(Declaration target,
            ICommand findReferencesCommand, 
            ICommand navigateCommand,
            ICommand closeCommand)
        {
            Target = target;
            FindReferencesCommand = findReferencesCommand;
            NavigateCommand = navigateCommand;
            CloseCommand = closeCommand;
        }

        public int MaxLines => 1000; // todo make this configurable?

        private Declaration _target;
        public Declaration Target
        {
            get => _target;
            set
            {
                if (_target != value)
                {
                    _target = value;
                    OnPropertyChanged();
                    SetPeekBody();
                }
            }
        }

        private void SetPeekBody()
        {
            if (Target == null)
            {
                Body = string.Empty;
                return;
            }

            ParserRuleContext context;
            if (Target.Context.Parent is VBAParser.ModuleBodyElementContext member)
            {
                context = member;
            }
            else if(Target.Context.TryGetAncestor<VBAParser.ModuleDeclarationsElementContext>(out var declaration))
            {
                context = declaration;
            }
            else
            {
                context = Target.Context;
            }

            var body = (context?.GetText() ?? string.Empty).Split('\n');
            var ellipsis = body?.Length > MaxLines ? "\n…" : string.Empty;

            Body = string.Join("\n", body.Take(MaxLines)) + ellipsis;
        }

        private string _body;
        public string Body
        {
            get => _body;
            set
            {
                if (_body != value)
                {
                    _body = value;
                    OnPropertyChanged();
                }
            }
        }

        public ICommand CloseCommand { get; }
        public ICommand FindReferencesCommand { get; }
        public ICommand NavigateCommand { get; }

        private object _navigateCommandParameter;
        public object NavigateCommandParameter
        {
            get => _navigateCommandParameter;
            set
            {
                if (_navigateCommandParameter != value)
                {
                    _navigateCommandParameter = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}