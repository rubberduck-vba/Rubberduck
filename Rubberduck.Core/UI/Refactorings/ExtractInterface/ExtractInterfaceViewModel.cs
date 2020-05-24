using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        public ExtractInterfaceViewModel(ExtractInterfaceModel model) : base(model)
        {
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
