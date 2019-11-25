using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public class EncapsulateFieldViewModel : RefactoringViewModelBase<EncapsulateFieldModel>
    {
        public RubberduckParserState State { get; }


        public EncapsulateFieldViewModel(EncapsulateFieldModel model, RubberduckParserState state/*, IIndenter indenter*/) : base(model)
        {
            State = state;

            IsLetSelected = true;
            PropertyName = model[model.TargetDeclaration].PropertyName;

            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));

            IsReadOnlyCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => RefreshPreview());
            EncapsulateFlagCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => RefreshPreview());
            PropertyOrFieldNameChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => RefreshPreview());
            BackingFieldNameChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => RefreshPreview());
        }

        public IEncapsulatedFieldViewData SelectedValue { set; get; }

        public Declaration TargetDeclaration
        {
            get => Model.TargetDeclaration;
            set
            {
                Model.TargetDeclaration = value;
                PropertyName = Model[Model.TargetDeclaration].PropertyName;
            }
        }

        public ObservableCollection<IEncapsulatedFieldViewData> EncapsulationFields
        {
            get
            {
                var flaggedFields = Model.EncapsulationFields.Where(efd => efd.EncapsulateFlag)
                    .OrderBy(efd => efd.Declaration.IdentifierName);

                var orderedFields = Model.EncapsulationFields.Except(flaggedFields)
                    .OrderBy(efd => efd.Declaration.IdentifierName);

                var viewableFields = new ObservableCollection<IEncapsulatedFieldViewData>();
                foreach (var efd in flaggedFields.Concat(orderedFields))
                {
                    viewableFields.Add(new ViewableEncapsulatedField(efd));
                }
                //TODO: Trying to reset the scroll to the top using SelectedValue is not working...Remove or fix 
                SelectedValue = viewableFields.FirstOrDefault();
                return viewableFields;
            }
        }

        public string LatestEdit
        {
            set
            {
                RefreshPreview();
            }
        }


        public bool EncapsulateAsUDT
        {
            get => Model.EncapsulateWithUDT;
            set
            {
                Model.EncapsulateWithUDT = value;
                RefreshPreview();
                OnPropertyChanged(nameof(EncapsulateAsUDT_TypeIdentifier));
                OnPropertyChanged(nameof(EncapsulateAsUDT_FieldName));
            }
        }

        public string EncapsulateAsUDT_TypeIdentifier
        {
            get => Model.EncapsulateWithUDT_TypeIdentifier;
            set
            {
                Model.EncapsulateWithUDT_TypeIdentifier = value;
                RefreshPreview();
            }
        }

        public string EncapsulateAsUDT_FieldName
        {
            get => Model.EncapsulateWithUDT_FieldName;
            set
            {
                Model.EncapsulateWithUDT_FieldName = value;
                RefreshPreview();
            }
        }

        public bool TargetsHaveValidEncapsulationSettings
        {
            get
            {
                return Model.EncapsulationFields.Where(efd => efd.EncapsulateFlag)
                    .Any(ff => ff.HasValidEncapsulationAttributes == false);
            }
        }


        public bool CanHaveLet => Model.CanImplementLet;
        public bool CanHaveSet => Model.CanImplementSet;

        public bool IsLetSelected
        {
            get => Model.ImplementLetSetterType;
            set
            {
                Model.ImplementLetSetterType = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public bool IsSetSelected
        {
            get => Model.ImplementSetSetterType;
            set
            {
                Model.ImplementSetSetterType = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public string PropertyName
        {
            get => Model.PropertyName;
            set
            {
                Model.PropertyName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidPropertyName));
                OnPropertyChanged(nameof(HasValidNames));
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public string ParameterName
        {
            get => Model.ParameterName;
            set
            {
                Model.ParameterName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidParameterName));
                OnPropertyChanged(nameof(HasValidNames));
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public IEncapsulateFieldNamesValidator RefactoringValidator { set; get; }

        public bool IsValidPropertyName
        {
            get
            {
                var encapsulatedField = Model[TargetDeclaration];

                return encapsulatedField.Declaration != null
                        && VBAIdentifierValidator.IsValidIdentifier(encapsulatedField.PropertyName, DeclarationType.Variable)
                        && !encapsulatedField.PropertyName.Equals(encapsulatedField.EncapsulationAttributes.NewFieldName, StringComparison.InvariantCultureIgnoreCase)
                        && !encapsulatedField.PropertyName.Equals(ParameterName, StringComparison.InvariantCultureIgnoreCase);
            }
        }

        public bool IsValidParameterName
        {
            get
            {
                var encapsulatedField = Model[TargetDeclaration];

                return encapsulatedField.Declaration != null
                        && VBAIdentifierValidator.IsValidIdentifier(encapsulatedField.PropertyName, DeclarationType.Variable)
                        && !encapsulatedField.EncapsulationAttributes.ParameterName.Equals(encapsulatedField.Declaration.IdentifierName, StringComparison.InvariantCultureIgnoreCase)
                        && !encapsulatedField.EncapsulationAttributes.ParameterName.Equals(encapsulatedField.EncapsulationAttributes.PropertyName, StringComparison.InvariantCultureIgnoreCase);
            }
        }


        public bool HasValidNames => IsValidPropertyName && IsValidParameterName;

        public string PropertyPreview
        {
            get
            {
                return Model.NewContent.AsSingleTextBlock;
            }
        }

        public CommandBase SelectAllCommand { get; }
        public CommandBase DeselectAllCommand { get; }
        private void ToggleSelection(bool value)
        {
            foreach (var item in EncapsulationFields)
            {
                item.EncapsulateFlag = value;
            }
            RefreshPreview();
        }

        public CommandBase IsReadOnlyCommand { get; }
        public CommandBase EncapsulateFlagCommand { get; }
        public CommandBase PropertyOrFieldNameChangeCommand { get; }
        public CommandBase BackingFieldNameChangeCommand { get; }

        private void RefreshPreview()
        {
            OnPropertyChanged(nameof(EncapsulationFields));
            OnPropertyChanged(nameof(PropertyPreview));
            OnPropertyChanged(nameof(SelectedValue));
        }

        //public event EventHandler<bool> ExpansionStateChanged;
        //private void OnExpansionStateChanged(bool value) => ExpansionStateChanged?.Invoke(this, value);
    }
}