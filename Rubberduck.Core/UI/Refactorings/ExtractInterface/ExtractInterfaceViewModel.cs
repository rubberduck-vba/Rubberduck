using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings
{
    public class ExtractInterfaceViewModel : RefactoringViewModelBase<ExtractInterfaceModel>
    {
        private readonly IConflictSession _conflictSession;
        private readonly IConflictDetectionModuleDeclarationProxy _newModuleProxy;

        public ExtractInterfaceViewModel(ExtractInterfaceModel model, IConflictSessionFactory conflictSessionFactory) : base(model)
        {
            _conflictSession = conflictSessionFactory.Create();
            _newModuleProxy = _conflictSession.ProxyCreator.CreateNewModule(Model.TargetDeclaration.ProjectId, VBEditor.SafeComWrappers.ComponentType.ClassModule, Model.InterfaceName);

            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));

            ComponentNames = Model.DeclarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .Where(moduleDeclaration => moduleDeclaration.ProjectId == Model.TargetDeclaration.ProjectId)
                .Select(module => module.ComponentName)
                .ToList();
        }

        public ObservableCollection<InterfaceMember> Members
        {
            get => Model.Members;
            set
            {
                Model.Members = value;
                OnPropertyChanged();
            }
        }

        public List<string> ComponentNames { get; set; }

        public string InterfaceName
        {
            get => Model.InterfaceName;
            set
            {
                Model.InterfaceName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidInterfaceName));
            }
        }

        public bool IsValidInterfaceName
        {
            get
            {
                _newModuleProxy.IdentifierName = InterfaceName;
                return VBAIdentifierValidator.IsValidIdentifier(_newModuleProxy.IdentifierName, DeclarationType.Module)
                        && _newModuleProxy.IdentifierName.Length > 1
                        && !_conflictSession.NewEntityConflictDetector.HasConflictingName(_newModuleProxy, out _);
            }
        }

        private void ToggleSelection(bool value)
        {
            foreach (var item in Members)
            {
                item.IsSelected = value;
            }
            OnPropertyChanged(nameof(Members));
        }

        public CommandBase SelectAllCommand { get; }
        public CommandBase DeselectAllCommand { get; }
    }
}
