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
        private class MasterDetailSelectionManager
        {
            private const string _neverATargeID = "_Never_a_TargetID_";
            private bool _detailFieldIsFlagged;

            public MasterDetailSelectionManager(string targetID)
            {
                SelectionTargetID = targetID;
                DetailField = null;
                _detailFieldIsFlagged = false;
            }


            private IEncapsulatedFieldViewData _detailField;
            public IEncapsulatedFieldViewData DetailField
            {
                set
                {
                    _detailField = value;
                    _detailFieldIsFlagged = _detailField?.EncapsulateFlag ?? false;
                }
                get => _detailField;
            }

            public string SelectionTargetID { set; get; }

            public bool DetailUpdateRequired
            {
                get
                {
                    if (DetailField is null)
                    {
                        return true;
                    }

                    if (_detailFieldIsFlagged != DetailField.EncapsulateFlag)
                    {
                        _detailFieldIsFlagged = !_detailFieldIsFlagged;
                        return true;
                    }
                    return SelectionTargetID != (DetailField?.TargetID ?? _neverATargeID);
                }
            }
        }

        private MasterDetailSelectionManager _masterDetail;
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

            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();
            var selectedField = model.SelectedFieldCandidates.FirstOrDefault();

            _masterDetail = new MasterDetailSelectionManager(selectedField.TargetID);
            if (selectedField != null)
            {
                _masterDetail.DetailField = EncapsulationFields.Where(ef => ef.EncapsulateFlag).Single();
            }

            ManageEncapsulationFlagsAndSelectedItem(selectedField);

            RefreshValidationResults();
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

        public IEncapsulatedFieldViewData SelectedField
        {
            set
            {
                if (value is null) { return; }

                _masterDetail.SelectionTargetID = value.TargetID;
                OnPropertyChanged(nameof(SelectedField));
                if (_masterDetail.DetailUpdateRequired)
                {
                    _masterDetail.DetailField = SelectedField;
                    UpdateDetailForSelection();
                }

                OnPropertyChanged(nameof(PropertyPreview));
            }

            get => EncapsulationFields.FirstOrDefault(f => f.TargetID.Equals(_masterDetail.SelectionTargetID)); // _selectedField;
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
            OnPropertyChanged(nameof(PropertyPreview));
            OnPropertyChanged(nameof(EncapsulationFields));
            OnPropertyChanged(nameof(ValidationErrorMessage));
        }

        public string PropertyName
        {
            set
            {
                if (SelectedField is null || value is null) { return; }

                _masterDetail.DetailField.PropertyName = value;
                UpdateDetailForSelection();
            }

            get => _masterDetail.DetailField?.PropertyName ?? SelectedField?.PropertyName ?? string.Empty; // _propertyName;
        }

        public bool SelectedFieldIsNotFlagged
            => !(_masterDetail.DetailField?.EncapsulateFlag ?? false);

        public bool SelectedFieldIsPrivateUDT
            => (_masterDetail.DetailField?.IsPrivateUserDefinedType ?? false);

        public bool SelectedFieldHasEditablePropertyName => !SelectedFieldIsPrivateUDT;

        public bool EnableReadOnlyOption 
            => !(_masterDetail.DetailField?.IsRequiredToBeReadOnly ?? false);

        public string GroupBoxHeaderContent 
            => $"{_masterDetail.DetailField?.TargetID ?? string.Empty} {EncapsulateFieldResources.GroupBoxHeaderSuffix} ";

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
                _masterDetail.DetailField.IsReadOnly = value;
                OnPropertyChanged(nameof(IsReadOnly));
            }
            get => _masterDetail.DetailField?.IsReadOnly ?? SelectedField?.IsReadOnly ?? false;
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

        private bool _hasValidNames;
        public bool HasValidNames => _hasValidNames;

        private bool _selectionHasValidEncapsulationAttributes;
        public bool SelectionHasValidEncapsulationAttributes => _selectionHasValidEncapsulationAttributes;

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
            if (_failedValidationResults.TryGetValue(_masterDetail.SelectionTargetID, out var errorMsg))
            {
                _validationErrorMessage = errorMsg;
                _selectionHasValidEncapsulationAttributes = false;
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
            if (value)
            {
                SelectedField = _lastCheckedBoxes.FirstOrDefault();
            }
            else
            {
                _masterDetail.DetailField = null;
            }
            ReloadListAndPreview();
            RefreshValidationResults();
            UpdateDetailForSelection();
        }

        public CommandBase IsReadOnlyCommand { get; }
        public CommandBase PropertyChangeCommand { get; }

        public CommandBase EncapsulateFlagChangeCommand { get; }
        public CommandBase ReadOnlyChangeCommand { get; }

        public string Caption
            => EncapsulateFieldResources.Caption;

        public string InstructionText
            => EncapsulateFieldResources.InstructionText;

        public string Preview
            => EncapsulateFieldResources.Preview;

        public string TitleText
            => EncapsulateFieldResources.TitleText;

        public string PrivateUDTPropertyText
            => EncapsulateFieldResources.PrivateUDTPropertyText;


        private void ChangeIsReadOnlyFlag(object param)
        {
            if (SelectedField is null) { return; }

            _masterDetail.DetailField.IsReadOnly = SelectedField.IsReadOnly;
            OnPropertyChanged(nameof(IsReadOnly));
            OnPropertyChanged(nameof(PropertyPreview));
        }

        private List<IEncapsulatedFieldViewData> _lastCheckedBoxes = new List<IEncapsulatedFieldViewData>();
        private void ManageEncapsulationFlagsAndSelectedItem(object param)
        {
            var selected = _lastCheckedBoxes.FirstOrDefault();
            if (_lastCheckedBoxes.Count == EncapsulationFields.Where(f => f.EncapsulateFlag).Count())
            {
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

            _masterDetail.SelectionTargetID = selected.TargetID;
            OnPropertyChanged(nameof(SelectedField));
            if (_masterDetail.DetailUpdateRequired)
            {
                _masterDetail.DetailField = SelectedField;
                UpdateDetailForSelection();
            }
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