using System.Runtime.InteropServices;
using NLog;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.AddRemoveReferences;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class AddRemoveReferencesCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IAddRemoveReferencesPresenterFactory _factory;
        private readonly IReferenceReconciler _reconciler;

        public AddRemoveReferencesCommand(IVBE vbe, 
            RubberduckParserState state, 
            IAddRemoveReferencesPresenterFactory factory,
            IReferenceReconciler reconciler) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
            _factory = factory;
            _reconciler = reconciler;
        }

        protected override void OnExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            var declaration = parameter is CodeExplorerItemViewModel explorerItem
                ? GetDeclaration(explorerItem)
                : GetDeclaration();

            if (!(Declaration.GetProjectParent(declaration) is ProjectDeclaration project))
            {
                return;
            }

            var dialog = _factory.Create(project);
            var model = dialog.Show();
            if (model is null)
            {
                return;
            }

            _reconciler.ReconcileReferences(model);
            _state.OnParseRequested(this);
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            if (parameter is CodeExplorerItemViewModel explorerNode)
            {
                return GetDeclaration(explorerNode) is ProjectDeclaration;
            }

            using (var project = _vbe.ActiveVBProject)
            {
                return !(project is null);
            }
        }

        private Declaration GetDeclaration(CodeExplorerItemViewModel node)
        {
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return (node as ICodeExplorerDeclarationViewModel)?.Declaration;
        }

        private Declaration GetDeclaration()
        {
            using (var project = _vbe.ActiveVBProject)
            {
                if (project is null || project.IsWrappingNullReference)
                {
                    return null;
                }
                return _state.DeclarationFinder.FindProject(project.Name);
            }           
        }
    }
}
