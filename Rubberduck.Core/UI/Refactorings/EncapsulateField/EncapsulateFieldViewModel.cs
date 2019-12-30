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
            private const string _neverATargetID = "_Never_a_TargetID_";
            private bool _detailFieldIsFlagged;

            public MasterDetailSelectionManager(IEncapsulateFieldCandidate selected)
                : this(selected?.TargetID)
            {
                if (selected != null)
                {
                    DetailField = new ViewableEncapsulatedField(selected);
                }
            }

            private MasterDetailSelectionManager(string targetID)
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

            private string _selectionTargetID;
            public string SelectionTargetID
            {
                set => _selectionTargetID = value;
                get => _selectionTargetID ?? _neverATargetID;
            }

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
                    return SelectionTargetID != DetailField?.TargetID;
                }
            }
        }

        private MasterDetailSelectionManager _masterDetail;
        public RubberduckParserState State { get; }

        public EncapsulateFieldViewModel(EncapsulateFieldModel model, RubberduckParserState state) : base(model)
        {
            State = state;

            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));

            IsReadOnlyCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdatePreview());
            PropertyChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdatePreview());
            EncapsulateFlagChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ManageEncapsulationFlagsAndSelectedItem);
            ReadOnlyChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ChangeIsReadOnlyFlag);

            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();

            _masterDetail = new MasterDetailSelectionManager(model.SelectedFieldCandidates.SingleOrDefault());

            ManageEncapsulationFlagsAndSelectedItem();

            RefreshValidationResults();
        }

        public ObservableCollection<IEncapsulatedFieldViewData> EncapsulationFields
        {
            get
            {
                var viewableFields = new ObservableCollection<IEncapsulatedFieldViewData>();

                var orderedFields = Model.EncapsulationCandidates.Where(ec => !(_selectedObjectStateUDT?.IsEncapsulateFieldCandidate(ec) ?? false))
                                            .OrderBy(efd => efd.Declaration.Selection).ToList();

                foreach (var efd in orderedFields)
                {
                    viewableFields.Add(new ViewableEncapsulatedField(efd));
                }

                return viewableFields;
            }
        }

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
                Model.StateUDTField = _selectedObjectStateUDT;
                SetObjectStateUDT();
            }
        }

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

                OnPropertyChanged(nameof(PropertiesPreview));
            }

            get => EncapsulationFields.FirstOrDefault(f => f.TargetID.Equals(_masterDetail.SelectionTargetID));
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

        public string PropertyName
        {
            set
            {
                if (SelectedField is null || value is null) { return; }

                _masterDetail.DetailField.PropertyName = value;
                UpdateDetailForSelection();
            }

            get => _masterDetail.DetailField?.PropertyName ?? SelectedField?.PropertyName ?? string.Empty;
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
            }
            get => _masterDetail.DetailField?.IsReadOnly ?? SelectedField?.IsReadOnly ?? false;
        }

        public bool ConvertFieldsToUDTMembers
        {
            get => Model.ConvertFieldsToUDTMembers;
            set
            {
                Model.ConvertFieldsToUDTMembers = value;
                ReloadListAndPreview();
                RefreshValidationResults();
                UpdateDetailForSelection();
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

        public string PropertiesPreview => Model.PreviewRefactoring();

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
            _masterDetail.SelectionTargetID = selected?.TargetID ?? null;
            OnPropertyChanged(nameof(SelectedField));
            if (_masterDetail.DetailUpdateRequired)
            {
                _masterDetail.DetailField = SelectedField;
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