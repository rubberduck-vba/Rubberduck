using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command;
using Rubberduck.UI.Refactorings.ExtractInterface;

namespace Rubberduck.UI.Refactorings
{
    public class ExtractInterfaceViewModel : RefactoringViewModelBase<ExtractInterfaceModel>
    {
        public ExtractInterfaceViewModel(ExtractInterfaceModel model) : base(model)
        {
            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));

            ComponentNames = Model.State.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .Where(moduleDeclaration => moduleDeclaration.ProjectId == Model.TargetDeclaration.ProjectId)
                .Select(module => module.ComponentName)
                .ToList();
            _members = Model.Members.Select(m => m.ToViewModel()).ToList();
            UpdateModelMembers();
        }

        private void UpdateModelMembers()
        {
            Model.Members = _members.Where(m => m.IsSelected).Select(vm => vm.ToModel()).ToList();
        }

        private List<InterfaceMemberViewModel> _members;
        public List<InterfaceMemberViewModel> Members
        {
            get => _members;
            set
            {
                _members = value;
                UpdateModelMembers();
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

        private void ToggleSelection(bool value)
        {
            foreach (var item in Members)
            {
                item.IsSelected = value;
            }
            UpdateModelMembers();
        }

        public CommandBase SelectAllCommand { get; }
        public CommandBase DeselectAllCommand { get; }
    }
}
