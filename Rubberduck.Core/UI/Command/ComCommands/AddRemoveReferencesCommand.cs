using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.AddRemoveReferences;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    [ComVisible(false)]
    public class AddRemoveReferencesCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IAddRemoveReferencesPresenterFactory _factory;
        private readonly IReferenceReconciler _reconciler;

        public AddRemoveReferencesCommand(IVBE vbe, 
            RubberduckParserState state, 
            IAddRemoveReferencesPresenterFactory factory,
            IReferenceReconciler reconciler,
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _state = state;
            _factory = factory;
            _reconciler = reconciler;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            if (parameter is CodeExplorerItemViewModel explorerNode)
            {
                return explorerNode.Declaration is ProjectDeclaration;
            }

            using (var project = _vbe.ActiveVBProject)
            {
                return !(project is null);
            }
        }

        protected override void OnExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            var declaration = parameter is CodeExplorerItemViewModel explorerItem
                ? explorerItem.Declaration
                : GetDeclaration();

            if (!(Declaration.GetProjectParent(declaration) is ProjectDeclaration project))
            {
                return;
            }

            var dialog = _factory.Create(project);

            var model = dialog?.Show();
            if (model is null)
            {
                return;
            }

            _reconciler.ReconcileReferences(model);
            _state.OnParseRequested(this);
        }

        private Declaration GetDeclaration()
        {
            using (var project = _vbe.ActiveVBProject)
            {
                if (project is null || project.IsWrappingNullReference)
                {
                    return null;
                }

                return _state.DeclarationFinder.Projects.OfType<ProjectDeclaration>()
                    .FirstOrDefault(declaration => project.ProjectId.Equals(declaration.ProjectId));
            }           
        }
    }
}
