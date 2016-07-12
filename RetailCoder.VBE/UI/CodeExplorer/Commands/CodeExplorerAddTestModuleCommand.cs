using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerAddTestModuleCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly NewUnitTestModuleCommand _newUnitTestModuleCommand;

        public CodeExplorerAddTestModuleCommand(VBE vbe, NewUnitTestModuleCommand newUnitTestModuleCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _newUnitTestModuleCommand = newUnitTestModuleCommand;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            try
            {
                return GetDeclaration(parameter) != null || _vbe.VBProjects.Count == 1;
            }
            catch (COMException)
            {
                return false;
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            if (parameter != null)
            {
                _newUnitTestModuleCommand.NewUnitTestModule(GetDeclaration(parameter).Project);
            }
            else
            {
                _newUnitTestModuleCommand.NewUnitTestModule(_vbe.VBProjects.Item(1));
            }
        }

        private Declaration GetDeclaration(object parameter)
        {
            var node = parameter as CodeExplorerItemViewModel;
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return node == null ? null : ((ICodeExplorerDeclarationViewModel)node).Declaration;
        }
    }
}
