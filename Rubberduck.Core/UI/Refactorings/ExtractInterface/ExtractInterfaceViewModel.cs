using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings
{
    public class ExtractInterfaceViewModel : RefactoringViewModelBase<ExtractInterfaceModel>
    {
        private static ObservableCollection<KeyValuePair<ExtractInterfaceImplementationOption, string>> _implementationOptionsExceptReplace;

        private static ObservableCollection<KeyValuePair<ExtractInterfaceImplementationOption, string>> _implementationOptionsAll;

        public ExtractInterfaceViewModel(ExtractInterfaceModel model) : base(model)
        {
            _implementationOptionsExceptReplace  = new ObservableCollection<KeyValuePair<ExtractInterfaceImplementationOption, string>>()
            {
                new KeyValuePair<ExtractInterfaceImplementationOption, string>(ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface, Resources.RubberduckUI.ExtractInterface_OptionForwardToInterfaceMembers),
                new KeyValuePair<ExtractInterfaceImplementationOption, string>(ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers, Resources.RubberduckUI.ExtractInterface_OptionForwardToObjectMembers),
                new KeyValuePair<ExtractInterfaceImplementationOption, string>(ExtractInterfaceImplementationOption.NoInterfaceImplementation, Resources.RubberduckUI.ExtractInterface_OptionAddEmptyImplementation),
            };

            _implementationOptionsAll = new ObservableCollection<KeyValuePair<ExtractInterfaceImplementationOption, string>>(_implementationOptionsExceptReplace)
            {
                new KeyValuePair<ExtractInterfaceImplementationOption, string>(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface, Resources.RubberduckUI.ExtractInterface_OptionReplaceMembersWithInterfaceMembers),
            };

            ResetImplementationOptions();

            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));

            ComponentNames = Model.DeclarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .Where(moduleDeclaration => moduleDeclaration.ProjectId == Model.TargetDeclaration.ProjectId)
                .Select(module => module.ComponentName)
                .ToList();

            foreach (var member in Model.Members)
            {
                member.PropertyChanged += HandleMemberSelectionChanged;
            }
        }

        public ObservableCollection<InterfaceMember> Members
        {
            get => Model.Members;
            set
            {
                Model.Members = value;

                ResetImplementationOptions();

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

        public ObservableCollection<KeyValuePair<ExtractInterfaceImplementationOption, string>> ImplementationOptions { set; get; }

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

        private void HandleMemberSelectionChanged(Object sender, PropertyChangedEventArgs args)
        {
            if (args.PropertyName == nameof(InterfaceMember.IsSelected))
            {
                ResetImplementationOptions();
            }
        }

        private void ResetImplementationOptions()
        {
            var selectedMemberHasExtReferences = Model.SelectedMembers.SelectMany(m => m.Member.References)
                .Any(rf => rf.QualifiedModuleName != rf.Declaration.QualifiedModuleName);

            ImplementationOptions = selectedMemberHasExtReferences
                ? _implementationOptionsExceptReplace
                : _implementationOptionsAll;

            if (selectedMemberHasExtReferences
                && ImplementationOption == ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)
            {
                ImplementationOption = ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface;
            }

            OnPropertyChanged(nameof(ImplementationOptions));
        }
    }
}
