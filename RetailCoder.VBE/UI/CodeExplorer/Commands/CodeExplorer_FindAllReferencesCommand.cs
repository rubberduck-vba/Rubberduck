using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_FindAllReferencesCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly FindAllReferencesCommand _findAllReferences;

        public CodeExplorer_FindAllReferencesCommand(RubberduckParserState state, FindAllReferencesCommand findAllReferences) : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
            _findAllReferences = findAllReferences;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _state.Status == ParserState.Ready &&
                parameter != null &&
                !(parameter is CodeExplorerCustomFolderViewModel);
        }

        protected override void ExecuteImpl(object parameter)
        {
            _findAllReferences.Execute(((CodeExplorerItemViewModel) parameter).GetSelectedDeclaration());
        }
    }
}
