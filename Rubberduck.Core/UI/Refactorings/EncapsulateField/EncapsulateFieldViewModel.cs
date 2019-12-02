using System.Collections.ObjectModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateField.Strategies;
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

        public ObservableCollection<IEncapsulatedFieldViewData> EncapsulationFields
        {
            get
            {
                var flaggedFields = Model.FlaggedEncapsulationFields
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

        public IEncapsulatedFieldViewData SelectedValue { set; get; }

        public bool HideEncapsulateAsUDTFields => !EncapsulateAsUDT;

        public bool EncapsulateAsUDT
        {
            get => Model.EncapsulateWithUDT;
            set
            {
                Model.EncapsulateWithUDT = value;
                UpdatePreview();
                OnPropertyChanged(nameof(HideEncapsulateAsUDTFields));
                OnPropertyChanged(nameof(EncapsulateAsUDT_TypeIdentifier));
                OnPropertyChanged(nameof(EncapsulateAsUDT_FieldName));
            }
        }

        public string EncapsulateAsUDT_TypeIdentifier
        {
            get
            {
                if (Model.EncapsulationStrategy is IEncapsulateWithBackingUserDefinedType udtStrategy)
                {
                    return udtStrategy.StateEncapsulationField.AsTypeName;
                }
                return string.Empty;
            }
            set
            {
                if (Model.EncapsulationStrategy is IEncapsulateWithBackingUserDefinedType udtStrategy)
                {
                    udtStrategy.StateEncapsulationField.EncapsulationAttributes.AsTypeName = value;
                }
                UpdatePreview();
            }
        }

        public string EncapsulateAsUDT_FieldName
        {
            get
            {
                if (Model.EncapsulationStrategy is IEncapsulateWithBackingUserDefinedType udtStrategy)
                {
                    return udtStrategy.StateEncapsulationField.NewFieldName;
                }
                return string.Empty;
            }
            set
            {
                if (Model.EncapsulationStrategy is IEncapsulateWithBackingUserDefinedType udtStrategy)
                {
                    udtStrategy.StateEncapsulationField.EncapsulationAttributes.NewFieldName = value;
                }
                UpdatePreview();
            }
        }

        public bool TargetsHaveValidEncapsulationSettings 
            => Model.EncapsulationFields.Where(efd => efd.EncapsulateFlag)
                    .Any(ff => !ff.HasValidEncapsulationAttributes);

        public IEncapsulateFieldNamesValidator RefactoringValidator { set; get; }

        //TODO: hook the validation scheme backup
        public bool HasValidNames => true;

        public string PropertyPreview => Model.PreviewRefactoring();

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