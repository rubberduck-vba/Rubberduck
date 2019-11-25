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

            SelectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(true));
            DeselectAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ToggleSelection(false));

            IsReadOnlyCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ReloadPreview());
            EncapsulateFlagCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ReloadPreview());
            PropertyChangeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdatePreview());
        }

        public IEncapsulatedFieldViewData SelectedValue { set; get; }

        //public Declaration TargetDeclaration
        //{
        //    get => Model.TargetDeclaration;
        //    set
        //    {
        //        Model.TargetDeclaration = value;
        //        //PropertyName = Model[Model.TargetDeclaration].PropertyName;
        //    }
        //}

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
                UpdatePreview();
            }
        }


        public bool EncapsulateAsUDT
        {
            get => Model.EncapsulateWithUDT;
            set
            {
                Model.EncapsulateWithUDT = value;
                UpdatePreview();
                //OnPropertyChanged(nameof(EncapsulateAsUDT_TypeIdentifier));
                //OnPropertyChanged(nameof(EncapsulateAsUDT_FieldName));
            }
        }

        public string EncapsulateAsUDT_TypeIdentifier
        {
            get => Model.EncapsulateWithUDT_TypeIdentifier;
            set
            {
                Model.EncapsulateWithUDT_TypeIdentifier = value;
                UpdatePreview();
            }
        }

        public string EncapsulateAsUDT_FieldName
        {
            get => Model.EncapsulateWithUDT_FieldName;
            set
            {
                Model.EncapsulateWithUDT_FieldName = value;
                UpdatePreview();
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


        //public bool CanHaveLet => Model.CanImplementLet;
        //public bool CanHaveSet => Model.CanImplementSet;

        //public bool IsLetSelected
        //{
        //    get => Model.ImplementLetSetterType;
        //    set
        //    {
        //        Model.ImplementLetSetterType = value;
        //        OnPropertyChanged();
        //        OnPropertyChanged(nameof(PropertyPreview));
        //    }
        //}

        //public bool IsSetSelected
        //{
        //    get => Model.ImplementSetSetterType;
        //    set
        //    {
        //        Model.ImplementSetSetterType = value;
        //        OnPropertyChanged();
        //        OnPropertyChanged(nameof(PropertyPreview));
        //    }
        //}

        //public string PropertyName
        //{
        //    get => Model.PropertyName;
        //    set
        //    {
        //        Model.PropertyName = value;
        //        OnPropertyChanged();
        //        OnPropertyChanged(nameof(IsValidPropertyName));
        //        OnPropertyChanged(nameof(HasValidNames));
        //        OnPropertyChanged(nameof(PropertyPreview));
        //    }
        //}

        //public string ParameterName
        //{
        //    get => Model.ParameterName;
        //    set
        //    {
        //        Model.ParameterName = value;
        //        OnPropertyChanged();
        //        OnPropertyChanged(nameof(IsValidParameterName));
        //        OnPropertyChanged(nameof(HasValidNames));
        //        OnPropertyChanged(nameof(PropertyPreview));
        //    }
        //}

        public IEncapsulateFieldNamesValidator RefactoringValidator { set; get; }

        //public bool IsValidPropertyName
        //{
        //    get
        //    {
        //        var encapsulatedField = Model[TargetDeclaration];

        //        return encapsulatedField.Declaration != null
        //                && VBAIdentifierValidator.IsValidIdentifier(encapsulatedField.PropertyName, DeclarationType.Variable);
        //                //&& !encapsulatedField.PropertyName.Equals(encapsulatedField.EncapsulationAttributes.NewFieldName, StringComparison.InvariantCultureIgnoreCase)
        //                //&& !encapsulatedField.PropertyName.Equals(ParameterName, StringComparison.InvariantCultureIgnoreCase);
        //    }
        //}

        //public bool IsValidParameterName
        //{
        //    get
        //    {
        //        var encapsulatedField = Model[TargetDeclaration];

        //        return encapsulatedField.Declaration != null
        //                && VBAIdentifierValidator.IsValidIdentifier(encapsulatedField.PropertyName, DeclarationType.Variable);
        //                //&& !encapsulatedField.EncapsulationAttributes.ParameterName.Equals(encapsulatedField.Declaration.IdentifierName, StringComparison.InvariantCultureIgnoreCase)
        //                //&& !encapsulatedField.EncapsulationAttributes.ParameterName.Equals(encapsulatedField.EncapsulationAttributes.PropertyName, StringComparison.InvariantCultureIgnoreCase);
        //    }
        //}

        //TODO: hook the validation scheme backup
        public bool HasValidNames => true; // IsValidPropertyName; // && IsValidParameterName;

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
            ReloadPreview();
        }

        public CommandBase IsReadOnlyCommand { get; }
        public CommandBase EncapsulateFlagCommand { get; }
        public CommandBase PropertyChangeCommand { get; }

        private void UpdatePreview()
        {
            OnPropertyChanged(nameof(PropertyPreview));
        }

        private void ReloadPreview()
        {
            OnPropertyChanged(nameof(EncapsulationFields));
            OnPropertyChanged(nameof(PropertyPreview));
        }
    }
}