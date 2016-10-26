using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class FindAllReferencesCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly Command.FindAllReferencesCommand _findAllReferences;

        public FindAllReferencesCommand(RubberduckParserState state, Command.FindAllReferencesCommand findAllReferences) : base(LogManager.GetCurrentClassLogger())
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
