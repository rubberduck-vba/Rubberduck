using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings
{
    public class ExtractInterfaceViewModel : ViewModelBase
    {
        public ExtractInterfaceViewModel()
        {
            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());
            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));
        }

        private List<InterfaceMember> _members;
        public List<InterfaceMember> Members
        {
            get { return _members; }
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
            get { return _interfaceName; }
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

        public event EventHandler<DialogResult> OnWindowClosed;
        private void DialogCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);
        private void DialogOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);

        private void ToggleSelection(bool value)
        {
            foreach (var item in Members)
            {
                item.IsSelected = value;
            }
        }

        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
        public CommandBase SelectAllCommand { get; }
        public CommandBase DeselectAllCommand { get; }
    }
}
