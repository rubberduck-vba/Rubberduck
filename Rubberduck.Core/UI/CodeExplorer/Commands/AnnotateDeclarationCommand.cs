using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AnnotateDeclarationCommand : CodeExplorerCommandBase
    {
        private static readonly Type[] ApplicableNodes =
        {
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private readonly RubberduckParserState _state;

        private readonly IRefactoringAction<AnnotateDeclarationModel> _annotateAction;
        private readonly IRefactoringFailureNotifier _failureNotifier;
        private readonly IRefactoringUserInteraction<AnnotateDeclarationModel> _userInteraction;

        public AnnotateDeclarationCommand(
            AnnotateDeclarationRefactoringAction annotateAction,
            AnnotateDeclarationFailedNotifier failureNotifier,
            RefactoringUserInteraction<IAnnotateDeclarationPresenter, AnnotateDeclarationModel> userInteraction,
            IVbeEvents vbeEvents, 
            RubberduckParserState state) 
            : base(vbeEvents)
        {
            _annotateAction = annotateAction;
            _failureNotifier = failureNotifier;
            _userInteraction = userInteraction;
            _state = state;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public override IEnumerable<Type> ApplicableNodeTypes => new[] { typeof(System.ValueTuple<IAnnotation, ICodeExplorerNode>) };

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            if (parameter is System.ValueTuple<IAnnotation, ICodeExplorerNode> data)
            {
                var (annotation, node) = data;
                return EvaluateCanExecute(annotation, node);
            }

            return false;
        }

        private bool EvaluateCanExecute(IAnnotation annotation, ICodeExplorerNode node)
        {
            var target = node?.Declaration;

            if (target == null 
                || annotation == null
                || !CanExecuteForNode(node))
            {
                return false;
            }

            if (!annotation.AllowMultiple 
                && target.Annotations.Any(pta => pta.Annotation.Equals(annotation)))
            {
                return false;
            }

            var targetType = target.DeclarationType;

            switch (annotation.Target)
            {
                case AnnotationTarget.Member:
                    return targetType.HasFlag(DeclarationType.Member)
                           && targetType != DeclarationType.LibraryFunction
                           && targetType != DeclarationType.LibraryProcedure;
                case AnnotationTarget.Module:
                    return targetType.HasFlag(DeclarationType.Module);
                case AnnotationTarget.Variable:
                    return targetType.HasFlag(DeclarationType.Variable)
                           || targetType.HasFlag(DeclarationType.Constant);
                case AnnotationTarget.General:
                    return true;
                case AnnotationTarget.Identifier:
                    return false;
                default:
                    return false;
            }
        }

        public bool CanExecuteForNode(ICodeExplorerNode node)
        {
            if (!ApplicableNodes.Contains(node.GetType())
                || !(node is CodeExplorerItemViewModel)
                || node.Declaration == null)
            {
                return false;
            }

            var target = node.Declaration;
            var targetType = target.DeclarationType;

            if (!targetType.HasFlag(DeclarationType.Module)
                && !targetType.HasFlag(DeclarationType.Variable)
                && !targetType.HasFlag(DeclarationType.Constant)
                && !targetType.HasFlag(DeclarationType.Member)
                || targetType == DeclarationType.LibraryFunction
                || targetType == DeclarationType.LibraryProcedure)
            {
                return false;
            }

            return !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        protected override void OnExecute(object parameter)
        {
            if (!(parameter is System.ValueTuple<IAnnotation, ICodeExplorerNode> data))
            {
                return;
            }

            var (annotation, node) = data;
            var target = node?.Declaration;
            try
            {
                var model = ModelFromParameter(annotation, target);
                if (!annotation.AllowedArguments.HasValue 
                    || annotation.AllowedArguments.Value > 0
                    || annotation is IAttributeAnnotation)
                {
                    model = _userInteraction.UserModifiedModel(model);
                }

                _annotateAction.Refactor(model);
            }
            catch (RefactoringAbortedException)
            {}
            catch (RefactoringException exception)
            {
                _failureNotifier.Notify(exception);
            }
        }

        private AnnotateDeclarationModel ModelFromParameter(IAnnotation annotation, Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            return new AnnotateDeclarationModel(target, annotation);
        }
    }
}