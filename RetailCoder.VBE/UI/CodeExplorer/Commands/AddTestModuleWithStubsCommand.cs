using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class AddTestModuleWithStubsCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly Command.AddTestModuleCommand _newUnitTestModuleCommand;

        public AddTestModuleWithStubsCommand(IVBE vbe, Command.AddTestModuleCommand newUnitTestModuleCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _newUnitTestModuleCommand = newUnitTestModuleCommand;
        }

        protected override bool EvaluateCanExecute(object parameter) => parameter is CodeExplorerComponentViewModel;

        protected override void OnExecute(object parameter)
        {
            if (parameter != null)
            {
                _newUnitTestModuleCommand.Execute(GetDeclaration(parameter));
            }
            else
            {
                _newUnitTestModuleCommand.Execute(_vbe.ActiveVBProject);
            }
        }

        private Declaration GetDeclaration(object parameter)
        {
            var node = parameter as CodeExplorerItemViewModel;
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return ((ICodeExplorerDeclarationViewModel)node)?.Declaration;
        }
    }
}
