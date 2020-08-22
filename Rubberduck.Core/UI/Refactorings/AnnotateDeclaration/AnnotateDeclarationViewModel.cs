using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.UI.Command;
using Rubberduck.UI.WPF;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    internal class AnnotateDeclarationViewModel : RefactoringViewModelBase<AnnotateDeclarationModel>
    {
        private readonly IAnnotationArgumentViewModelFactory _argumentFactory;

        public AnnotateDeclarationViewModel(
            AnnotateDeclarationModel model,
            IEnumerable<IAnnotation> annotations,
            IAnnotationArgumentViewModelFactory argumentFactory
        )
            : base(model)
        {
            _argumentFactory = argumentFactory;
            ApplicableAnnotations = AnnotationsForDeclaration(model.Target, annotations);
            AnnotationArguments = new ObservableViewModelCollection<IAnnotationArgumentViewModel>();
            RefreshAnnotationArguments(Model.Annotation);

            AddAnnotationArgument = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteAddArgument, parameter => CanAddArgument);
            RemoveAnnotationArgument = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRemoveArgument, parameter => CanRemoveArgument);
        }

        public IReadOnlyList<IAnnotation> ApplicableAnnotations { get; }

        private static IReadOnlyList<IAnnotation> AnnotationsForDeclaration(Declaration declaration, IEnumerable<IAnnotation> annotations)
        {
            return AnnotationsForDeclarationType(declaration.DeclarationType, annotations)
                .Where(annotation => (annotation.AllowMultiple 
                                        || !declaration.Annotations.Any(pta => annotation.Equals(pta.Annotation)))
                                     && (declaration.DeclarationType.HasFlag(DeclarationType.Module) 
                                        || declaration.AttributesPassContext != null 
                                        || !(annotation is IAttributeAnnotation)))
                .ToList();
        }

        private static IEnumerable<IAnnotation> AnnotationsForDeclarationType(DeclarationType declarationType, IEnumerable<IAnnotation> annotations)
        {
            if (declarationType.HasFlag(DeclarationType.Module))
            {
                return annotations.Where(annotation => annotation.Target.HasFlag(AnnotationTarget.Module));
            }

            if (declarationType.HasFlag(DeclarationType.Member) 
                && declarationType != DeclarationType.LibraryProcedure
                && declarationType != DeclarationType.LibraryFunction)
            {
                return annotations.Where(annotation => annotation.Target.HasFlag(AnnotationTarget.Member));
            }

            if (declarationType.HasFlag(DeclarationType.Variable))
            {
                return annotations.Where(annotation => annotation.Target.HasFlag(AnnotationTarget.Variable));
            }

            return Enumerable.Empty<IAnnotation>();
        }

        public IAnnotation Annotation
        {
            get => Model.Annotation;
            set
            {
                if (value == null && Model.Annotation == null || (value?.Equals(Model.Annotation) ?? false))
                {
                    return;
                }

                Model.Annotation = value;
                RefreshAnnotationArguments(value);

                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidAnnotation));
                OnPropertyChanged(nameof(ShowAdjustAttributeOption));
            }
        }

        public ObservableViewModelCollection<IAnnotationArgumentViewModel> AnnotationArguments { get; }

        public bool AdjustAttribute
        {
            get => Model.AdjustAttribute;
            set
            {
                if (value == Model.AdjustAttribute)
                {
                    return;
                }

                Model.AdjustAttribute = value;

                OnPropertyChanged();
            }
        }

        public bool ShowAdjustAttributeOption => Model?.Annotation is IAttributeAnnotation;

        private void RefreshAnnotationArguments(IAnnotation annotation)
        {
            AnnotationArguments.Clear();
            var newArguments = InitialAnnotationArguments(annotation);

            foreach (var argument in newArguments)
            {
                AnnotationArguments.Add(argument);
            }
        }

        private IEnumerable<IAnnotationArgumentViewModel> InitialAnnotationArguments(IAnnotation annotation)
        {
            return annotation != null
                    ? annotation.AllowedArgumentTypes
                        .Select(InitialArgumentViewModel)
                        .Take(annotation.RequiredArguments)
                : Enumerable.Empty<AnnotationArgumentViewModel>();
        }

        private IAnnotationArgumentViewModel InitialArgumentViewModel(AnnotationArgumentType argumentType)
        {
            var argumentModel = _argumentFactory.Create(argumentType, string.Empty);
            argumentModel.ErrorsChanged += ArgumentErrorStateChanged;
            return argumentModel;
        }

        private void ArgumentErrorStateChanged(object requestor, DataErrorsChangedEventArgs e)
        {
            OnPropertyChanged(nameof(IsValidAnnotation));
        }

        public CommandBase AddAnnotationArgument { get; }
        public CommandBase RemoveAnnotationArgument { get; }

        private bool CanAddArgument => Annotation != null
                                       && (!Annotation.AllowedArguments.HasValue
                                           || AnnotationArguments.Count < Annotation.AllowedArguments.Value);

        private bool CanRemoveArgument => Annotation != null 
                                          && AnnotationArguments.Count > Model.Annotation.RequiredArguments;

        private void ExecuteAddArgument(object parameter)
        {
            if (Annotation == null)
            {
                return;
            }

            var argumentType = Annotation.AllowedArgumentTypes.Last();
            var newArgument = InitialArgumentViewModel(argumentType);
            AnnotationArguments.Add(newArgument);
        }

        private void ExecuteRemoveArgument(object parameter)
        {
            if (!AnnotationArguments.Any())
            {
                return;
            }

            AnnotationArguments.RemoveAt(AnnotationArguments.Count - 1);
        }

        protected override void DialogOk()
        {
            Model.Arguments = AnnotationArguments
                .Select(viewModel => viewModel.Model)
                .ToList();

            base.DialogOk();
        }

        public bool IsValidAnnotation => Annotation != null 
                                          && AnnotationArguments.All(argument => !argument.HasErrors);
    }
}