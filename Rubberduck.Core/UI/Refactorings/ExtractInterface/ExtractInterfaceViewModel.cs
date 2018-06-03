using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command;
using Rubberduck.UI.Refactorings.ExtractInterface;

namespace Rubberduck.UI.Refactorings
{
    internal class ExtractInterfaceViewModel : RefactoringViewModelBase<ExtractInterfaceModel>
    {
        public ExtractInterfaceViewModel(ExtractInterfaceModel model) : base(model)
        {
            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));
        }

        private List<InterfaceMemberViewModel> _members;
        public List<InterfaceMemberViewModel> Members
        {
            get => _members;
            set
            {
                _members = value;
                OnPropertyChanged();
            }
        }

        public List<string> ComponentNames { get; set; }

        private string _interfaceName;
        public string InterfaceName
        {
            get => _interfaceName;
            set
            {
                _interfaceName = value;
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
        }

        public CommandBase SelectAllCommand { get; }
        public CommandBase DeselectAllCommand { get; }
    }
}
