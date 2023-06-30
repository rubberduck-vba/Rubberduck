using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings // Todo: Correct Namespace
{
    public class ExtractInterfaceViewModel : RefactoringViewModelBase<ExtractInterfaceModel>
    {
        private static ObservableCollection<KeyValuePair<ExtractInterfaceImplementationOption, string>> _implementationOptions;

        public ExtractInterfaceViewModel(ExtractInterfaceModel model) : base(model)
        {
            _implementationOptions  = new ObservableCollection<KeyValuePair<ExtractInterfaceImplementationOption, string>>()
            {
                new KeyValuePair<ExtractInterfaceImplementationOption, string>(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface, RefactoringsUI.ExtractInterface_OptionForwardToInterfaceMembers),
                new KeyValuePair<ExtractInterfaceImplementationOption, string>(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers, RefactoringsUI.ExtractInterface_OptionForwardToObjectMembers),
                new KeyValuePair<ExtractInterfaceImplementationOption, string>(ExtractInterfaceImplementationOption.NoInterfaceImplementation, RefactoringsUI.ExtractInterface_OptionAddEmptyImplementation),
                new KeyValuePair<ExtractInterfaceImplementationOption, string>(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface, RefactoringsUI.ExtractInterface_OptionReplaceMembersWithInterfaceMembers)
            };

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
                if (!VBAIdentifierValidator.IsValidIdentifier(InterfaceName, DeclarationType.ClassModule))
                {
                    return false;
                }

                return !Model.ConflictFinder.IsConflictingModuleName(InterfaceName);
            }
        }

        public bool CanChooseInterfaceInstancing => Model.ImplementingClassInstancing != ClassInstancing.Public;

        public IEnumerable<ClassInstancing> ClassInstances => Enum.GetValues(typeof(ClassInstancing)).Cast<ClassInstancing>();

        public ClassInstancing InterfaceInstancing
        {
            get => Model.InterfaceInstancing;

            set
            {
                if (value == Model.InterfaceInstancing)
                {
                    return;
                }

                Model.InterfaceInstancing = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<KeyValuePair<ExtractInterfaceImplementationOption, string>> ImplementationOptions => _implementationOptions;

        public ExtractInterfaceImplementationOption ImplementationOption
        {
            get => Model.ImplementationOption;

            set
            {
                if (Model.ImplementationOption == value)
                {
                    return;
                }

                Model.ImplementationOption = value;
                OnPropertyChanged();
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
