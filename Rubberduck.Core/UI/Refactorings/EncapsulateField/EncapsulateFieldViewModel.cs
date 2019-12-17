using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public class EncapsulateFieldViewModel : RefactoringViewModelBase<EncapsulateFieldModel>
    {
        public RubberduckParserState State { get; }

        public EncapsulateFieldViewModel(EncapsulateFieldModel model, RubberduckParserState state/*, IIndenter indenter*/) : base(model)
        {
            State = state;

            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));

            IsReadOnlyCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdatePreview());
            PropertyChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdatePreview());
            EncapsulateFlagChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ManageEncapsulationFlagsAndSelectedItem);
            ReadOnlyChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ChangeIsReadOnlyFlag);

            //SelectedField = EncapsulationFields.FirstOrDefault(ef => ef.EncapsulateFlag);
            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();
            ManageEncapsulationFlagsAndSelectedItem(null);
        }

        public ObservableCollection<IEncapsulatedFieldViewData> EncapsulationFields
        {
            get
            {
                var viewableFields = new ObservableCollection<IEncapsulatedFieldViewData>();

                var orderedFields = Model.EncapsulationCandidates
                    .OrderBy(efd => efd.Declaration.Selection);

                foreach (var efd in orderedFields)
                {
                    viewableFields.Add(new ViewableEncapsulatedField(efd));
                }

                return viewableFields;
            }
        }

        private string _selectedTargetID;
        public IEncapsulatedFieldViewData SelectedField
        {
            set
            {
                if (value is null) { return; }

                _selectedTargetID = value.TargetID;
                var _selectedField = value;
                PropertyName = _selectedField.IsPrivateUserDefinedType
                    ? "Encapsulates UDT Members"
                    : _selectedField.PropertyName;

                IsReadOnly = _selectedField.IsReadOnly;

                OnPropertyChanged(nameof(EnableReadOnlyOption));
                OnPropertyChanged(nameof(EnablePropertyNameEditor));
                OnPropertyChanged(nameof(HideGroupBox));
                OnPropertyChanged(nameof(GroupBoxHeaderContent));
                OnPropertyChanged(nameof(SelectedField));
                OnPropertyChanged(nameof(HasValidNames));
                OnPropertyChanged(nameof(BackingField));
                OnPropertyChanged(nameof(PropertyPreview));
            }

            get => EncapsulationFields.FirstOrDefault(f => f.TargetID.Equals(_selectedTargetID)); // _selectedField;
        }

        public bool HideGroupBox
            => !(SelectedField?.EncapsulateFlag ?? true);

        public bool EnablePropertyNameEditor
            => !(SelectedField?.IsPrivateUserDefinedType ?? false);

        public bool EnableReadOnlyOption 
            => !(SelectedField?.IsRequiredToBeReadOnly ?? false);

        public string GroupBoxHeaderContent 
            => $"Encapsulation Property for Field: {SelectedField?.TargetID ?? string.Empty}";

        public string BackingField 
            => $"Backing Field : {SelectedField?.NewFieldName ?? string.Empty}";

        string _propertyName;
        public string PropertyName
        {
            set
            {
                if (SelectedField is null) { return; }

                _propertyName = value;
                SelectedField.PropertyName = value;
                OnPropertyChanged(nameof(PropertyName));
                OnPropertyChanged(nameof(BackingField));
                OnPropertyChanged(nameof(HasValidNames));
                OnPropertyChanged(nameof(HasValidEncapsulationAttributes));
                OnPropertyChanged(nameof(PropertyPreview));
                OnPropertyChanged(nameof(EncapsulationFields));
            }
            get => _propertyName;
        }

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                return SelectedField?.HasValidEncapsulationAttributes ?? false;
            }
        }

        private string _fieldDescriptor;
        public string FieldDescriptor
        {
            set
            {
                _fieldDescriptor = value;
                OnPropertyChanged(nameof(FieldDescriptor));
            }
            get => _fieldDescriptor;
        }

        private string _targetID;
        public string TargetID
        {
            set
            {
                _targetID = value;
            }
            get => _targetID;
        }

        private bool _isReadOnly;
        public bool IsReadOnly
        {
            set
            {
                _isReadOnly = value;
                OnPropertyChanged(nameof(IsReadOnly));
            }
            get => _isReadOnly;
        }

        public bool EncapsulateAsUDT
        {
            get => Model.EncapsulateWithUDT;
            set
            {
                Model.EncapsulateWithUDT = value;
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public bool HasValidNames
        {
            get
            {
                return EncapsulationFields.All(ef => ef.HasValidEncapsulationAttributes);
            }
        }

        public string PropertyPreview => Model.PreviewRefactoring();

        public CommandBase SelectAllCommand { get; }
        public CommandBase DeselectAllCommand { get; }
        private void ToggleSelection(bool value)
        {
            _lastCheckedBoxes.Clear();
            foreach (var item in EncapsulationFields)
            {
                item.EncapsulateFlag = value;
            }
            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();
            ReloadListAndPreview();
        }

        public CommandBase IsReadOnlyCommand { get; }
        public CommandBase PropertyChangeCommand { get; }

        public CommandBase EncapsulateFlagChangeCommand { get; }
        public CommandBase ReadOnlyChangeCommand { get; }


        private void ChangeIsReadOnlyFlag(object param)
        {
            if (SelectedField is null) { return; }

            SelectedField.IsReadOnly = !SelectedField.IsReadOnly;
            OnPropertyChanged(nameof(IsReadOnly));
            OnPropertyChanged(nameof(PropertyPreview));
        }

        private List<IEncapsulatedFieldViewData> _lastCheckedBoxes;

        private void ManageEncapsulationFlagsAndSelectedItem(object param)
        {
            var selected = _lastCheckedBoxes.FirstOrDefault();
            if (_lastCheckedBoxes.Count == EncapsulationFields.Count)
            {
                SelectedField = selected;
                return;
            }

            var nowChecked = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();
            var beforeChecked = _lastCheckedBoxes.ToList();

            nowChecked.RemoveAll(c => _lastCheckedBoxes.Contains(c));
            beforeChecked.RemoveAll(c => EncapsulationFields.Where(ec => ec.EncapsulateFlag).Select(nc => nc).Contains(c)); //.TargetID));
            if (nowChecked.Any())
            {
                selected = nowChecked.First();
            }
            else if (beforeChecked.Any())
            {
                selected = beforeChecked.First();
            }
            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();
            SelectedField = selected;
        }

        private void UpdatePreview() 
            => OnPropertyChanged(nameof(PropertyPreview));

        private void ReloadListAndPreview()
        {
            OnPropertyChanged(nameof(EncapsulationFields));
            OnPropertyChanged(nameof(PropertyPreview));
        }
    }
}