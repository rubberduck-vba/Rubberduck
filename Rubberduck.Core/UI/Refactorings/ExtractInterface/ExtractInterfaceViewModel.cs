using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings
{
    public class ExtractInterfaceViewModel : RefactoringViewModelBase<ExtractInterfaceModel>
    {
        private Dictionary<ExtractInterfaceImplementationOption, string> _implementationOptions;

        public ExtractInterfaceViewModel(ExtractInterfaceModel model) : base(model)
        {
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
                OnPropertyChanged(nameof(ImplementationOption));
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
                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

                return !ComponentNames.Contains(InterfaceName)
                       && InterfaceName.Length > 1
                       && char.IsLetter(InterfaceName.FirstOrDefault())
                       && !tokenValues.Contains(InterfaceName, StringComparer.InvariantCultureIgnoreCase)
                       && !InterfaceName.Any(c => !char.IsLetterOrDigit(c) && c != '_');
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

        public IEnumerable<string> ImplementationOptions => _implementationOptions.Values;

        public string ImplementationOption
        {
            get => _implementationOptions[Model.ImplementationOption];

            set
            {
                if (_implementationOptions.TryGetValue(Model.ImplementationOption, out var currentValue)
                    && value == currentValue)
                {
                    return;
                }

                Model.ImplementationOption = _implementationOptions.Single(op => op.Value == value).Key;
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

        //TODO: Load Option descriptors from UI Resources
        //TODO: (In XAML) Set Implementation Options Group Box label from UI Resources
        private void ResetImplementationOptions()
        {
            _implementationOptions = new Dictionary<ExtractInterfaceImplementationOption, string>()
            {
                [ExtractInterfaceImplementationOption.ForwardInterfaceToObjectMembers] = "Forward Interface Member Calls to Object Members",
                [ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface] = "Forward Object Member Calls to Interface Members",
                [ExtractInterfaceImplementationOption.NoInterfaceImplementation] = "Add Implementation TODO comments",
            };

            if (!Model.SelectedMembers.SelectMany(m => m.Member.References).Any(rf => rf.QualifiedModuleName != rf.Declaration.QualifiedModuleName))
            {
                _implementationOptions.Add(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface, "Replace Members with Interface Members");
            }

            OnPropertyChanged(nameof(ImplementationOptions));

            if (!_implementationOptions.ContainsKey(ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)
                && Model.ImplementationOption == ExtractInterfaceImplementationOption.ReplaceObjectMembersWithInterface)
            {
                ImplementationOption = _implementationOptions[ExtractInterfaceImplementationOption.ForwardObjectMembersToInterface];
            }
        }
    }
}
