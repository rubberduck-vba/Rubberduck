using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractInterface;

namespace Rubberduck.UI.Refactorings.ExtractInterface
{
    internal class ExtractInterfacePresenter : RefactoringPresenterBase<ExtractInterfaceModel, ExtractInterfaceDialog, ExtractInterfaceView, ExtractInterfaceViewModel>, IExtractInterfacePresenter
    {
        public ExtractInterfacePresenter(ExtractInterfaceModel model,
            IRefactoringDialogFactory dialogFactory) : base(model, dialogFactory)
        {
            ViewModel = dialogFactory.CreateViewModel<ExtractInterfaceModel, ExtractInterfaceViewModel>(model);
        }

        public override ExtractInterfaceViewModel ViewModel { get; }

        public override ExtractInterfaceModel Show()
        {
            if (Model.TargetDeclaration == null)
            {
                return null;
            }

            ViewModel.ComponentNames = Model.State.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .Where(moduleDeclaration => moduleDeclaration.ProjectId == Model.TargetDeclaration.ProjectId)
                .Select(module => module.ComponentName)
                .ToList();
            ViewModel.InterfaceName = Model.InterfaceName;
            ViewModel.Members = Model.Members.Select(m => m.ToViewModel()).ToList();

            Show();
            if (DialogResult != RefactoringDialogResult.Execute)
            {
                return null;
            }

            Model.InterfaceName = ViewModel.InterfaceName;
            Model.Members = ViewModel.Members.Where(m => m.IsSelected).Select(vm => vm.ToModel()).ToList();
            return Model;
        }
    }
}
