using System.Globalization;
using System.Linq;
using System.Windows.Input;
using Antlr4.Runtime;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Resources;

namespace Rubberduck.UI.Controls
{
    public class PeekDefinitionViewModel : ViewModelBase
    {
        private readonly IParseTreeProvider _parseTrees;

        internal PeekDefinitionViewModel()
        {
            // default constructor for xaml designer
        }

        public PeekDefinitionViewModel(ICodeExplorerNode node,
            ICommand findReferencesCommand,
            ICommand navigateCommand,
            ICommand closeCommand,
            IParseTreeProvider parseTrees)
        :this(node.Declaration, findReferencesCommand, navigateCommand, closeCommand, parseTrees)
        {}

        public PeekDefinitionViewModel(Declaration target,
            ICommand findReferencesCommand, 
            ICommand navigateCommand,
            ICommand closeCommand,
            IParseTreeProvider parseTrees)
        {
            _parseTrees = parseTrees;
            FindReferencesCommand = findReferencesCommand;
            NavigateCommand = navigateCommand;
            CloseCommand = closeCommand;
            Target = target;
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

        public string DescriptionString => string.IsNullOrWhiteSpace(_target.DescriptionString)
            ? RubberduckUI.ResourceManager.GetString("PeekDefinition_DefaultDescription", CultureInfo.CurrentUICulture)
            : _target.DescriptionString;

        private void SetPeekBody()
        {
            if (Target == null)
            {
                Body = string.Empty;
                return;
            }

            ParserRuleContext context;
            if (Target?.Context?.Parent is VBAParser.ModuleBodyElementContext member)
            {
                context = member;
            }
            else if(Target?.Context != null)
            {
                context = Target.Context.TryGetAncestor<VBAParser.ModuleDeclarationsElementContext>(out var declaration) ? (ParserRuleContext)declaration :
                          Target.Context.TryGetAncestor<VBAParser.BlockStmtContext>(out var statement) ? statement : null;
            }
            else if (Target is ModuleDeclaration module)
            {
                context = _parseTrees.GetParseTree(module.QualifiedModuleName, CodeKind.CodePaneCode) as ParserRuleContext;
            }
            else
            {
                context = null;
            }

            var body = (context?.GetText() ?? string.Empty).Split('\n');
            if (Target is ModuleDeclaration)
            {
                // body ends with an <EOF> token; strip it.
                body = body.Take(body.Length - 1).ToArray();
            }
            var ellipsis = body.Length > MaxLines ? "\n…" : string.Empty;
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
    }
}