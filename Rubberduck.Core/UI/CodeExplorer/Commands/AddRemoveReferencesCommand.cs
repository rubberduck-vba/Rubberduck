using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using NLog;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.AddRemoveReferences;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using AddRemoveReferencesViewModel = Rubberduck.UI.AddRemoveReferences.AddRemoveReferencesViewModel;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddRemoveReferencesCommand : CommandBase
    {
        private readonly IVBE _vbe;

        public AddRemoveReferencesCommand(IVBE vbe) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
        }

        protected override void OnExecute(object parameter)
        {
            if (parameter is CodeExplorerItemViewModel vm)
            {
                var project = GetDeclaration(vm)?.Project;
                if (project == null)
                {
                    return; 
                }

                var refs = new List<ReferenceModel>();
                using (var references = project.References)
                {
                    var priority = 1;
                    foreach (var reference in references)
                    {
                        refs.Add(new ReferenceModel(reference, priority++));
                        reference.Dispose();
                    }
                }

                var finder = new RegisteredLibraryFinderService(true);
                refs.AddRange(finder.FindRegisteredLibraries());

                using (var dialog = new AddRemoveReferencesDialog(new AddRemoveReferencesViewModel(refs)))
                {
                    dialog.ShowDialog();
                }
                
            }


        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return GetDeclaration(parameter as CodeExplorerItemViewModel) is ProjectDeclaration;
        }

        private Declaration GetDeclaration(CodeExplorerItemViewModel node)
        {
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return (node as ICodeExplorerDeclarationViewModel)?.Declaration;
        }
    }
}
