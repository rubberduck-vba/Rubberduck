using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Resources;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public class EncapsulateFieldViewModel : RefactoringViewModelBase<EncapsulateFieldModel>
    {
        private MasterDetailSelectionManager _masterDetailManager;

        public EncapsulateFieldViewModel(EncapsulateFieldModel model) : base(model)
        {
            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));

            IsReadOnlyCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdatePreview());
            PropertyChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdatePreview());
            EncapsulateFlagChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ManageEncapsulationFlagsAndSelectedItem);
            ReadOnlyChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ChangeIsReadOnlyFlag);

            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();

            _masterDetailManager = new MasterDetailSelectionManager(model.SelectedFieldCandidates.SingleOrDefault());

            ManageEncapsulationFlagsAndSelectedItem();

            RefreshValidationResults();
        }

        public ObservableCollection<IEncapsulatedFieldViewData> EncapsulationFields
        {
            get
            {
                var viewableFields = new ObservableCollection<IEncapsulatedFieldViewData>();

                var orderedFields = Model.EncapsulationCandidates.OrderBy(efd => efd.Declaration.Selection).ToList();
                if (_selectedObjectStateUDT != null && _selectedObjectStateUDT.IsExistingDeclaration)
                {
                    orderedFields = Model.EncapsulationCandidates.Where(ec => !_selectedObjectStateUDT.FieldIdentifier.Equals(ec.IdentifierName))
                                                .OrderBy(efd => efd.Declaration.Selection).ToList();
                }
                IsEmptyList = orderedFields.Count() == 0;
                foreach (var efd in orderedFields)
                {
                    viewableFields.Add(new ViewableEncapsulatedField(efd));
                }

                return viewableFields;
            }
        }

        public bool IsEmptyList { set; get; }

        public ObservableCollection<IObjectStateUDT> UDTFields
        {
            get
            {
                var viewableFields = new ObservableCollection<IObjectStateUDT>();

                foreach (var state in Model.ObjectStateUDTCandidates)
                {
                    viewableFields.Add(state);
                }
                return viewableFields;
            }
        }

        public bool ShowStateUDTSelections
        {
            get
            {
                return Model.ObjectStateUDTCandidates.Count() > 1
                    && ConvertFieldsToUDTMembers;
            }
        }

        private IObjectStateUDT _selectedObjectStateUDT;
        public IObjectStateUDT SelectedObjectStateUDT
        {
            get
            {
                    _selectedObjectStateUDT = UDTFields.Where(f => f.IsSelected)
                        .SingleOrDefault() ?? UDTFields.FirstOrDefault();
                return _selectedObjectStateUDT;
            }
            set
            {
                _selectedObjectStateUDT = value;
                Model.ObjectStateUDTField = _selectedObjectStateUDT;
                SetObjectStateUDT();
            }
        }

        public IEncapsulatedFieldViewData SelectedField
        {
            set
            {
                if (value is null) { return; }

                _masterDetailManager.SelectionTargetID = value.TargetID;
                OnPropertyChanged(nameof(SelectedField));
                if (_masterDetailManager.DetailUpdateRequired)
                {
                    _masterDetailManager.DetailField = SelectedField;
                    UpdateDetailForSelection();
                }

                OnPropertyChanged(nameof(PropertiesPreview));
            }

            get => EncapsulationFields.FirstOrDefault(f => f.TargetID.Equals(_masterDetailManager.SelectionTargetID));
        }

        public string PropertyName
        {
            set
            {
                if (SelectedField is null || value is null) { return; }

                _masterDetailManager.DetailField.PropertyName = value;
                UpdateDetailForSelection();
            }

            get => _masterDetailManager.DetailField?.PropertyName ?? SelectedField?.PropertyName ?? string.Empty;
        }

        public bool SelectedFieldIsNotFlagged
            => !(_masterDetailManager.DetailField?.EncapsulateFlag ?? false);

        public bool SelectedFieldIsPrivateUDT
            => (_masterDetailManager.DetailField?.IsPrivateUserDefinedType ?? false);

        public bool SelectedFieldHasEditablePropertyName => !SelectedFieldIsPrivateUDT;

        public bool EnableReadOnlyOption 
            => !(_masterDetailManager.DetailField?.IsRequiredToBeReadOnly ?? false);

        public string GroupBoxHeaderContent
            => $"{_masterDetailManager.DetailField?.TargetID ?? string.Empty} {RubberduckUI.EncapsulateField_PropertyName} ";

        private string _validationErrorMessage;
        public string ValidationErrorMessage => _validationErrorMessage;

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

        public bool IsReadOnly
        {
            set
            {
                _masterDetailManager.DetailField.IsReadOnly = value;
            }
            get => _masterDetailManager.DetailField?.IsReadOnly ?? SelectedField?.IsReadOnly ?? false;
        }

        public bool ConvertFieldsToUDTMembers
        {
            get => Model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers;
            set
            {
                Model.EncapsulateFieldStrategy = value
                    ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                    : EncapsulateFieldStrategy.UseBackingFields;
                ReloadListAndPreview();
                RefreshValidationResults();
                UpdateDetailForSelection();
                OnPropertyChanged(nameof(ShowStateUDTSelections));
            }
        }

        private bool _hasValidNames;
        public bool HasValidNames => _hasValidNames;

        private bool _selectionHasValidEncapsulationAttributes;
        public bool SelectionHasValidEncapsulationAttributes => _selectionHasValidEncapsulationAttributes;

        public string PropertiesPreview => Model.PreviewRefactoring();

        public CommandBase SelectAllCommand { get; }

        public CommandBase DeselectAllCommand { get; }

        public CommandBase IsReadOnlyCommand { get; }

        public CommandBase PropertyChangeCommand { get; }

        public CommandBase EncapsulateFlagChangeCommand { get; }

        public CommandBase ReadOnlyChangeCommand { get; }

        private void SetObjectStateUDT()
        {
            foreach (var field in UDTFields)
            {
                field.IsSelected = _selectedObjectStateUDT == field;
            }
            OnPropertyChanged(nameof(SelectedObjectStateUDT));
            OnPropertyChanged(nameof(EncapsulationFields));
            OnPropertyChanged(nameof(PropertiesPreview));
        }

        private void ToggleSelection(bool value)
        {
            _lastCheckedBoxes.Clear();
            foreach (var item in EncapsulationFields)
            {
                item.EncapsulateFlag = value;
            }
            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();
            if (value)
            {
                SelectedField = _lastCheckedBoxes.FirstOrDefault();
            }
            else
            {
                _masterDetailManager.DetailField = null;
            }
            ReloadListAndPreview();
            RefreshValidationResults();
            UpdateDetailForSelection();
        }

        private Dictionary<string, string> _failedValidationResults = new Dictionary<string, string>();
        private void RefreshValidationResults()
        {
            _failedValidationResults.Clear();
            _hasValidNames = true;
            _selectionHasValidEncapsulationAttributes = true;
            _validationErrorMessage = string.Empty;

            foreach (var field in EncapsulationFields.Where(ef => ef.EncapsulateFlag))
            {
                if (!field.TryValidateEncapsulationAttributes(out var errorMessage))
                {
                    _failedValidationResults.Add(field.TargetID, errorMessage);
                }
            }

            _hasValidNames = !_failedValidationResults.Any();
            if (_failedValidationResults.TryGetValue(_masterDetailManager.SelectionTargetID, out var errorMsg))
            {
                _validationErrorMessage = errorMsg;
                _selectionHasValidEncapsulationAttributes = false;
            }
        }

        private void UpdateDetailForSelection()
        {

            RefreshValidationResults();

            OnPropertyChanged(nameof(SelectedFieldIsNotFlagged));
            OnPropertyChanged(nameof(GroupBoxHeaderContent));
            OnPropertyChanged(nameof(PropertyName));
            OnPropertyChanged(nameof(IsReadOnly));
            OnPropertyChanged(nameof(HasValidNames));
            OnPropertyChanged(nameof(EnableReadOnlyOption));
            OnPropertyChanged(nameof(SelectedFieldIsPrivateUDT));
            OnPropertyChanged(nameof(SelectedFieldHasEditablePropertyName));
            OnPropertyChanged(nameof(SelectionHasValidEncapsulationAttributes));
            OnPropertyChanged(nameof(PropertiesPreview));
            OnPropertyChanged(nameof(EncapsulationFields));
            OnPropertyChanged(nameof(ValidationErrorMessage));
        }

        private void ChangeIsReadOnlyFlag(object param)
        {
            if (SelectedField is null) { return; }

            _masterDetailManager.DetailField.IsReadOnly = SelectedField.IsReadOnly;
            OnPropertyChanged(nameof(IsReadOnly));
            OnPropertyChanged(nameof(PropertiesPreview));
        }

        private List<IEncapsulatedFieldViewData> _lastCheckedBoxes = new List<IEncapsulatedFieldViewData>();
        private void ManageEncapsulationFlagsAndSelectedItem(object param = null)
        {
            var selected = _lastCheckedBoxes.FirstOrDefault();
            if (_lastCheckedBoxes.Count == EncapsulationFields.Where(f => f.EncapsulateFlag).Count())
            {
                _lastCheckedBoxes = EncapsulationFields.Where(f => f.EncapsulateFlag).ToList();
                if (EncapsulationFields.Where(f => f.EncapsulateFlag).Count() == 0
                    && EncapsulationFields.Count() > 0)
                {
                    SetSelectedField(EncapsulationFields.First());
                    return;
                }
                SetSelectedField(_lastCheckedBoxes.First());
                return;
            }

            var nowChecked = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();
            var beforeChecked = _lastCheckedBoxes.ToList();

            nowChecked.RemoveAll(c => _lastCheckedBoxes.Contains(c));
            beforeChecked.RemoveAll(c => EncapsulationFields.Where(ec => ec.EncapsulateFlag).Select(nc => nc).Contains(c));
            if (nowChecked.Any())
            {
                selected = nowChecked.First();
            }
            else if (beforeChecked.Any())
            {
                selected = beforeChecked.First();
            }
            else
            {
                selected = null;
            }

            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();

            SetSelectedField(selected);
        }

        private void SetSelectedField(IEncapsulatedFieldViewData selected)
        {
            _masterDetailManager.SelectionTargetID = selected?.TargetID ?? null;
            OnPropertyChanged(nameof(SelectedField));
            if (_masterDetailManager.DetailUpdateRequired)
            {
                _masterDetailManager.DetailField = SelectedField;
                UpdateDetailForSelection();
            }
        }

        private void UpdatePreview() 
            => OnPropertyChanged(nameof(PropertiesPreview));

        private void ReloadListAndPreview()
        {
            OnPropertyChanged(nameof(EncapsulationFields));
            OnPropertyChanged(nameof(PropertiesPreview));
        }
    }
}