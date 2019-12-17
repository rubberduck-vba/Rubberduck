using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
//using Rubberduck.Refactorings.EncapsulateField.Strategies;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public class EncapsulateFieldViewModel : RefactoringViewModelBase<EncapsulateFieldModel>
    {
        public RubberduckParserState State { get; }

        public EncapsulateFieldViewModel(EncapsulateFieldModel model, RubberduckParserState state/*, IIndenter indenter*/) : base(model)
        {
            State = state;

            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();

            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));

            IsReadOnlyCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ReloadPreview());
            EncapsulateFlagCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ReloadPreview());
            PropertyChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdatePreview());
            EncapsulateFlagChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ManageEncapsulationFlagsAndSelectedItem);
            ReadOnlyChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ChangeIsReadOnlyFlag);

            SelectedField = EncapsulationFields.FirstOrDefault();
        }

        private ObservableCollection<IEncapsulatedFieldViewData> _viewableFields;
        public ObservableCollection<IEncapsulatedFieldViewData> EncapsulationFields
        {
            get
            {
                if (_viewableFields is null)
                {
                    _viewableFields = new ObservableCollection<IEncapsulatedFieldViewData>();

                    var orderedFields = Model.EncapsulationCandidates
                        .OrderBy(efd => efd.Declaration.Selection);

                    foreach (var efd in orderedFields)
                    {
                        _viewableFields.Add(new ViewableEncapsulatedField(efd));
                    }
                }
                return _viewableFields;
            }
            set
            {
                var tossOut = value;
            }
        }

        private IEncapsulatedFieldViewData _selectedField;
        public IEncapsulatedFieldViewData SelectedField
        {
            set
            {
                _selectedField = value;
                PropertyName = _selectedField.PropertyName;
                GroupBoxHeaderContent = "Property Name";
                IsReadOnly = _selectedField.IsReadOnly;
                OnPropertyChanged(nameof(SelectedField));
                OnPropertyChanged(nameof(PropertyPreview));
            }
            get => _selectedField;
        }

        string _groupBoxHeader;
        public string GroupBoxHeaderContent
        {
            set
            {
                _groupBoxHeader = value;
                OnPropertyChanged(nameof(GroupBoxHeaderContent));
            }
            get => _groupBoxHeader;
        }

        string _propertyName;
        public string PropertyName
        {
            set
            {
                _propertyName = value;
                SelectedField.PropertyName = value;
                OnPropertyChanged(nameof(PropertyName));
                OnPropertyChanged(nameof(HasValidNames));
                OnPropertyChanged(nameof(HasValidEncapsulationAttributes));
                OnPropertyChanged(nameof(PropertyPreview));
            }
            get => _propertyName;
        }

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                return SelectedField.HasValidEncapsulationAttributes;
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

        //public void ValidatePropertyName(string value)
        //{
        //    SelectedField.PropertyName = value;
        //    OnPropertyChanged(nameof(HasValidNames));
        //}

        //public bool HideEncapsulateAsUDTFields => !EncapsulateAsUDT;

        public bool EncapsulateAsUDT
        {
            get => Model.EncapsulateWithUDT;
            set
            {
                Model.EncapsulateWithUDT = value;
                OnPropertyChanged(nameof(PropertyPreview));
                //OnPropertyChanged(nameof(HideEncapsulateAsUDTFields));
                //OnPropertyChanged(nameof(EncapsulateAsUDT_TypeIdentifier));
                //OnPropertyChanged(nameof(EncapsulateAsUDT_FieldName));
            }
        }

        //public string EncapsulateAsUDT_TypeIdentifier
        //{
        //    get
        //    {
        //        if (Model.EncapsulateWithUDT)
        //        {
        //            return Model.StateUDTField.TypeIdentifier;
        //        }
        //        return string.Empty;
        //    }
        //    set
        //    {
        //        if (Model.EncapsulateWithUDT)
        //        {
        //            Model.StateUDTField.TypeIdentifier = value;
        //        }
        //        OnPropertyChanged(nameof(PropertyPreview));
        //    }
        //}

        //public string EncapsulateAsUDT_FieldName
        //{
        //    get
        //    {
        //        if (Model.EncapsulateWithUDT)
        //        {
        //            return Model.StateUDTField.FieldIdentifier;
        //        }
        //        return string.Empty;
        //    }
        //    set
        //    {
        //        if (Model.EncapsulateWithUDT)
        //        {
        //            Model.StateUDTField.FieldIdentifier = value;
        //        }
        //        OnPropertyChanged(nameof(PropertyPreview));
        //    }
        //}

        //public bool TargetsHaveValidEncapsulationSettings 
        //    => Model.EncapsulationCandidates.Where(efd => efd.EncapsulateFlag)
        //            .Any(ff => !ff.HasValidEncapsulationAttributes);

        //public IEncapsulateFieldValidator RefactoringValidator { set; get; }

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
            foreach (var item in _viewableFields)
            {
                item.EncapsulateFlag = value;
            }
            ReloadPreview();
        }

        public CommandBase IsReadOnlyCommand { get; }
        public CommandBase EncapsulateFlagCommand { get; }
        public CommandBase PropertyChangeCommand { get; }

        public CommandBase EncapsulateFlagChangeCommand { get; }
        public CommandBase ReadOnlyChangeCommand { get; }


        private void ChangeIsReadOnlyFlag(object param)
        {
                SelectedField.IsReadOnly = !SelectedField.IsReadOnly;
                OnPropertyChanged(nameof(IsReadOnly));
                OnPropertyChanged(nameof(PropertyPreview));
        }

        private List<IEncapsulatedFieldViewData> _lastCheckedBoxes;

        private void ManageEncapsulationFlagsAndSelectedItem(object param)
        {
                        
            var nowChecked = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();
            var beforeChecked = _lastCheckedBoxes.ToList();

            nowChecked.RemoveAll(c => _lastCheckedBoxes.Select(bc => bc.TargetID).Contains(c.TargetID));
            beforeChecked.RemoveAll(c => EncapsulationFields.Where(ec => ec.EncapsulateFlag).Select(nc => nc.TargetID).Contains(c.TargetID));
            if (nowChecked.Any())
            {
                SelectedField = nowChecked.Single();
            }
            else if (beforeChecked.Any())
            {
                SelectedField = beforeChecked.Single();
            }
            _lastCheckedBoxes = EncapsulationFields.Where(ef => ef.EncapsulateFlag).ToList();
        }

        private void UpdatePreview() 
            => OnPropertyChanged(nameof(PropertyPreview));

        private void ReloadPreview()
        {
            OnPropertyChanged(nameof(EncapsulationFields));
            OnPropertyChanged(nameof(PropertyPreview));
        }
    }
}