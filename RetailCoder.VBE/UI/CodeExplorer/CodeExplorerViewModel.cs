using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

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

            _refreshCommand = new DelegateCommand(ExecuteRefreshCommand);
        }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand { get { return _refreshCommand; } }

        private readonly ICommand _toggleSignaturesCommand;
        public ICommand ToggleSignaturesCommand { get { return _toggleSignaturesCommand; } }

        private readonly ICommand _toggleNamespacesCommand;
        public ICommand ToggleNamespacesCommand { get { return _toggleNamespacesCommand; } }

        private readonly ICommand _toggleFoldersCommand;
        public ICommand ToggleFoldersCommand { get { return _toggleFoldersCommand; } }

        private object _selectedItem;
        public object SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value; 
                OnPropertyChanged();
            }
        }

        private bool _isBusy;

        public bool IsBusy
        {
            get { return _isBusy; }
            set
            {
                _isBusy = value; 
                OnPropertyChanged();
                CanRefresh = !_isBusy;
            }
        }

        private bool _canRefresh = true;
        public bool CanRefresh
        {
            get { return true /*_canRefresh*/; }
            private set
            {
                _canRefresh = value;
                OnPropertyChanged();
            }
        }

        private IEnumerable<CodeExplorerProjectViewModel> _projects;
        public IEnumerable<CodeExplorerProjectViewModel> Projects
        {
            get { return _projects; }
            set
            {
                _projects = value; 
                OnPropertyChanged();
            }
        } 

        private void ParserState_StateChanged(object sender, ParserStateEventArgs e)
        {
            IsBusy = e.State == ParserState.Parsing;
            if (e.State != ParserState.Parsed)
            {
                return;
            }

            var userDeclarations = _state.AllUserDeclarations
                .GroupBy(declaration => declaration.Project)
                .ToList();

            Projects = userDeclarations.Select(grouping => 
                new CodeExplorerProjectViewModel(grouping.Single(declaration => declaration.DeclarationType == DeclarationType.Project), grouping));
        }

        private void ParserState_ModuleStateChanged(object sender, Parsing.ParseProgressEventArgs e)
        {
            // todo: figure out a way to handle error state.
            // the problem is that the _projects collection might not contain our failing module yet.
        }

        private void ExecuteRefreshCommand(object param)
        {
            _state.OnParseRequested();
        }
    }

    public class CodeExplorerProjectViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;
        private readonly IEnumerable<CodeExplorerComponentViewModel> _components;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.Class, 
            DeclarationType.Document, 
            DeclarationType.Module, 
            DeclarationType.UserForm, 
        };

        public CodeExplorerProjectViewModel(Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _declaration = declaration;
            _components = declarations.GroupBy(item => item.ComponentName)
                .SelectMany(grouping =>
                    grouping.Where(item => ComponentTypes.Contains(item.DeclarationType))
                        .Select(item => new CodeExplorerComponentViewModel(item, grouping)));
        }

        public bool IsProtected { get { return _declaration.Project.Protection == vbext_ProjectProtection.vbext_pp_locked; } }
        public IEnumerable<CodeExplorerComponentViewModel> Components { get { return _components; } }
    }

    public class CodeExplorerComponentViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;
        private readonly IEnumerable<CodeExplorerMemberViewModel> _members;

        private static readonly DeclarationType[] MemberTypes =
        {
            DeclarationType.Constant, 
            DeclarationType.Enumeration, 
            DeclarationType.Event, 
            DeclarationType.Function, 
            DeclarationType.LibraryFunction, 
            DeclarationType.LibraryProcedure, 
            DeclarationType.Procedure,
            DeclarationType.PropertyGet, 
            DeclarationType.PropertyLet, 
            DeclarationType.PropertySet, 
            DeclarationType.UserDefinedType, 
            DeclarationType.Variable, 
        };

        public CodeExplorerComponentViewModel(Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _declaration = declaration;
            _members = declarations.GroupBy(item => item.Scope)
                .SelectMany(grouping =>
                    grouping.Where(item => MemberTypes.Contains(item.DeclarationType))
                        .Select(item => new CodeExplorerMemberViewModel(item)));
        }

        private bool _isErrorState;
        public bool IsErrorState { get { return _isErrorState; } set { _isErrorState = value; OnPropertyChanged(); } }

        public bool IsTestModule
        {
            get
            {
                return _declaration.DeclarationType == DeclarationType.Module 
                    && _declaration.Annotations.Split('\n').Contains(Parsing.Grammar.Annotations.TestModule);
            }
        }

        public string Name { get { return _declaration.IdentifierName; } }

        public string Namespace
        {
            get
            {
                var result = _declaration.Annotations
                    .Split('\n')
                    .FirstOrDefault(annotation => annotation.StartsWith(Parsing.Grammar.Annotations.Namespace));
                
                if (result == null)
                {
                    return string.Empty;
                }

                // don't throw in getter, but if value.Length > 2 then anything after the 2nd space is ignored:
                var value = result.Split(' ');
                return value.Length == 1 ? string.Empty : value[1];
            }
        }

        private vbext_ComponentType ComponentType { get { return _declaration.QualifiedName.QualifiedModuleName.Component.Type; } }

        private static readonly IDictionary<vbext_ComponentType, DeclarationType> DeclarationTypes = new Dictionary<vbext_ComponentType, DeclarationType>
        {
            { vbext_ComponentType.vbext_ct_ClassModule, DeclarationType.Class },
            { vbext_ComponentType.vbext_ct_StdModule, DeclarationType.Module },
            { vbext_ComponentType.vbext_ct_Document, DeclarationType.Document },
            { vbext_ComponentType.vbext_ct_MSForm, DeclarationType.UserForm }
        };

        public DeclarationType DeclarationType
        {
            get
            {
                DeclarationType result;
                if (!DeclarationTypes.TryGetValue(ComponentType, out result))
                {
                    result = DeclarationType.Class;
                }

                return result;
            }
        }
    }

    public class CodeExplorerMemberViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;

        public CodeExplorerMemberViewModel(Declaration declaration)
        {
            _declaration = declaration;
        }

        public string Name { get { return _declaration.IdentifierName; } }
        //public string Signature { get { return _declaration.IdentifierName; } }

    }
}
