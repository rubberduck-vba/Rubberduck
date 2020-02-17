using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParameterRefactoringAction : RefactoringActionBase<ReorderParametersModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ReorderParameterRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager) 
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Refactor(ReorderParametersModel model)
        {
            if (!model.Parameters.Where((param, index) => param.Index != index).Any())
            {
                //This is not an error: the user chose to leave everything as-is.
                return;
            }

            base.Refactor(model);
        }

        protected override void Refactor(ReorderParametersModel model, IRewriteSession rewriteSession)
        {
            AdjustReferences(model, rewriteSession);
            AdjustSignatures(model, rewriteSession);

            // if we are reordering a property getter, check if we need to reorder a letter/setter too
            if (model.IsPropertyRefactoringWithGetter)
            {
                var setter = GetLetterOrSetter(model.TargetDeclaration, DeclarationType.PropertySet);
                if (setter != null)
                {
                    var setterModel = ModelForNewTarget(model, setter);
                    Refactor(setterModel, rewriteSession);
                }

                var letter = GetLetterOrSetter(model.TargetDeclaration, DeclarationType.PropertyLet);
                if (letter != null)
                {
                    var letterModel = ModelForNewTarget(model, letter);
                    Refactor(letterModel, rewriteSession);
                }
            }

            var eventImplementations = _declarationFinderProvider.DeclarationFinder
                .FindEventHandlers(model.TargetDeclaration);

            foreach (var eventHandler in eventImplementations)
            {
                var eventHandlerModel = ModelForNewTarget(model, eventHandler);
                AdjustReferences(eventHandlerModel, rewriteSession);
                AdjustSignatures(eventHandlerModel, rewriteSession);
            }

            var interfaceImplementations = _declarationFinderProvider.DeclarationFinder
                .FindInterfaceImplementationMembers(model.TargetDeclaration);

            foreach (var interfaceImplementation in interfaceImplementations)
            {
                var interfaceImplementationModel = ModelForNewTarget(model, interfaceImplementation);
                AdjustReferences(interfaceImplementationModel, rewriteSession);
                AdjustSignatures(interfaceImplementationModel, rewriteSession);
            }
        }

        private static void AdjustReferences(ReorderParametersModel model, IRewriteSession rewriteSession)
        {
            var parameterDeclarations = model.Parameters
                .Select(param => param.Declaration)
                .ToList();
            var argumentReferences = ArgumentReferencesByLocation(parameterDeclarations);

            foreach (var (module, moduleArgumentReferences) in argumentReferences)
            {
                AdjustReferences(model, module, moduleArgumentReferences, rewriteSession);
            }
        }

        private static Dictionary<QualifiedModuleName, Dictionary<Selection, List<ArgumentReference>>> ArgumentReferencesByLocation(ICollection<ParameterDeclaration> parameters)
        {
            return parameters
                .SelectMany(parameterDeclaration => parameterDeclaration.ArgumentReferences)
                .GroupBy(argumentReference => argumentReference.QualifiedModuleName)
                .ToDictionary(
                    grouping => grouping.Key,
                    grouping => grouping
                        .GroupBy(reference => reference.ArgumentListSelection)
                        .ToDictionary(group => group.Key, group => group.ToList()));
        }

        private static void AdjustReferences(
            ReorderParametersModel model,
            QualifiedModuleName module,
            Dictionary<Selection, List<ArgumentReference>> argumentReferences,
            IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(module);
            foreach (var (argumentListSelection, sameArgumentListReferences) in argumentReferences)
            {
                //This happens for (with) dictionary access expressions only,
                //which cannot be reordered anyway.
                if (argumentListSelection.Equals(Selection.Empty))
                {
                    continue;
                }

                AdjustReferences(model, sameArgumentListReferences, rewriter);
            }
        }

        private static void AdjustReferences(ReorderParametersModel model, IReadOnlyCollection<ArgumentReference> argumentReferences, IModuleRewriter rewriter)
        {
            if (!argumentReferences.Any())
            {
                return;
            }

            if (argumentReferences.Any(argReference => argReference.ArgumentType == ArgumentListArgumentType.Named))
            {
                var positionalArguments = argumentReferences
                    .Where(argReference => argReference.ArgumentType == ArgumentListArgumentType.Positional);
                MakeArgumentsNamed(positionalArguments, rewriter);

                var missingArguments = argumentReferences
                    .Where(argReference => argReference.ArgumentType == ArgumentListArgumentType.Missing)
                    .ToList();
                RemoveArguments(missingArguments, rewriter);

                return;
            }

            var argumentReferencesWithoutParamArrayReferences = argumentReferences.Where(reference =>
                !((ParameterDeclaration)reference.Declaration).IsParamArray)
                .ToList();

            var argumentsWithNewPosition = ArgumentsWithNewPosition(model, argumentReferencesWithoutParamArrayReferences);

            if (argumentReferencesWithoutParamArrayReferences.Count == argumentReferences.Count)
            {
                //If no parameters for a param array are provided, the reordering can cause trailing missing arguments, which have to be removed.
                argumentsWithNewPosition = RemoveMissingArgumentsTrailingAfterReorder(argumentsWithNewPosition, rewriter).ToList();
            }

            ReorderArguments(argumentsWithNewPosition, rewriter);
        }

        private static void MakeArgumentsNamed(IEnumerable<ArgumentReference> argumentReferences, IModuleRewriter rewriter)
        {
            foreach (var argumentReference in argumentReferences)
            {
                if (argumentReference.ArgumentType != ArgumentListArgumentType.Positional)
                {
                    continue;
                }

                MakePositionalArgumentNamed(argumentReference, rewriter);
            }
        }

        private static void MakePositionalArgumentNamed(ArgumentReference argumentReference, IModuleRewriter rewriter)
        {
            var parameterName = argumentReference.Declaration.IdentifierName;
            var insertionCode = $"{parameterName}:=";
            var argumentContext = argumentReference.Context;
            var insertionIndex = argumentContext.Start.TokenIndex;
            rewriter.InsertBefore(insertionIndex, insertionCode);
        }

        private static void RemoveArguments(IReadOnlyCollection<ArgumentReference> argumentReferences, IModuleRewriter rewriter)
        {
            if (!argumentReferences.Any())
            {
                return;
            }

            var argumentIndicesToRemove = argumentReferences.Select(argumentReference => argumentReference.ArgumentPosition);
            var argumentIndexRangesToRemove = IndexRanges(argumentIndicesToRemove);
            var argumentList = argumentReferences.First().ArgumentListContext;

            foreach (var (startIndex, stopIndex) in argumentIndexRangesToRemove)
            {
                RemoveArgumentRange(startIndex, stopIndex, argumentList, rewriter);
            }
        }

        private static IEnumerable<(int startIndex, int stopIndex)> IndexRanges(IEnumerable<int> indices)
        {
            var sortedIndices = indices.OrderBy(num => num).ToList();
            var ranges = new List<(int startIndex, int stopIndex)>();
            int startIndex = -10;
            int stopIndex = -10;
            foreach (var currentIndex in sortedIndices)
            {
                if (currentIndex == stopIndex + 1)
                {
                    stopIndex = currentIndex;
                }
                else
                {
                    if (startIndex >= 0)
                    {
                        ranges.Add((startIndex, stopIndex));
                    }

                    startIndex = currentIndex;
                    stopIndex = currentIndex;
                }
            }

            if (startIndex >= 0)
            {
                ranges.Add((startIndex, stopIndex));
            }

            return ranges;
        }

        private static void RemoveArgumentRange(
            int startArgumentIndex,
            int stopArgumentIndex,
            VBAParser.ArgumentListContext argumentList,
            IModuleRewriter rewriter)
        {
            var (startTokenIndex, stopTokenIndex) = TokenIndexRange(startArgumentIndex, stopArgumentIndex, argumentList.argument());
            rewriter.RemoveRange(startTokenIndex, stopTokenIndex);
        }

        private static (int startTokenIndex, int stopTokenIndex) TokenIndexRange(
            int startIndex,
            int stopIndex,
            IReadOnlyList<ParserRuleContext> contexts)
        {
            int startTokenIndex;
            int stopTokenIndex;

            if (stopIndex == contexts.Count - 1)
            {
                startTokenIndex = startIndex == 0
                    ? contexts[0].Start.TokenIndex
                    : contexts[startIndex - 1].Stop.TokenIndex + 1;
                stopTokenIndex = contexts[stopIndex].Stop.TokenIndex;
                return (startTokenIndex, stopTokenIndex);
            }

            startTokenIndex = contexts[startIndex].Start.TokenIndex;
            stopTokenIndex = contexts[stopIndex + 1].Start.TokenIndex - 1;
            return (startTokenIndex, stopTokenIndex);
        }

        private static IEnumerable<(ArgumentReference argumentReference, int newIndex)> RemoveMissingArgumentsTrailingAfterReorder(
            IEnumerable<(ArgumentReference reference, int Index)> argumentsWithNewPosition,
            IModuleRewriter rewriter)
        {
            var argumentsWithNewPositionOrderedBackwards = argumentsWithNewPosition
                .OrderByDescending(tuple => tuple.Index)
                .ToList();

            var numberOfTrailingMissingMembers = NumberOfTrailingMissingArguments(argumentsWithNewPositionOrderedBackwards);

            if (numberOfTrailingMissingMembers > 0)
            {
                RemoveTrailingArguments(argumentsWithNewPositionOrderedBackwards, numberOfTrailingMissingMembers, rewriter);
            }

            return argumentsWithNewPositionOrderedBackwards.Skip(numberOfTrailingMissingMembers);
        }

        private static List<(ArgumentReference reference, int)> ArgumentsWithNewPosition(ReorderParametersModel model, IReadOnlyCollection<ArgumentReference> argumentReferences)
        {
            var newIndex = NewIndicesOfParameterIndices(model);
            var argumentsWithNewPosition = argumentReferences
                .Select(reference => (reference, newIndex[reference.ArgumentPosition]))
                .ToList();
            return argumentsWithNewPosition;
        }

        private static int[] NewIndicesOfParameterIndices(ReorderParametersModel model)
        {
            var newIndices = new int[model.Parameters.Count];
            foreach (var (parameter, index) in model.Parameters.Select((parameter, index) => (parameter, index)))
            {
                newIndices[parameter.Index] = index;
            }

            return newIndices;
        }

        private static int NumberOfTrailingMissingArguments(IReadOnlyList<(ArgumentReference reference, int Index)> argumentsWithNewPositionOrderedBackwards)
        {
            var numberOfTrailingMissingMembers = 0;
            while (numberOfTrailingMissingMembers < argumentsWithNewPositionOrderedBackwards.Count
                   && argumentsWithNewPositionOrderedBackwards[numberOfTrailingMissingMembers].reference.ArgumentType ==
                   ArgumentListArgumentType.Missing)
            {
                numberOfTrailingMissingMembers++;
            }

            return numberOfTrailingMissingMembers;
        }

        private static void RemoveTrailingArguments(
            IReadOnlyList<(ArgumentReference reference, int Index)> argumentsWithNewPositionOrderedBackwards,
            int numberOfTrailingMissingMembers,
            IModuleRewriter rewriter)
        {
            var stopNewArgumentIndex = argumentsWithNewPositionOrderedBackwards.Count - 1;
            var startNewArgumentIndex = stopNewArgumentIndex - numberOfTrailingMissingMembers + 1;

            var argumentList = argumentsWithNewPositionOrderedBackwards[0].reference.ArgumentListContext;
            RemoveArgumentRange(startNewArgumentIndex, stopNewArgumentIndex, argumentList, rewriter);
        }

        private static void ReorderArguments(
            IReadOnlyList<(ArgumentReference argumentReference, int newIndex)> argumentsWithNewPosition,
            IModuleRewriter rewriter)
        {
            if (!argumentsWithNewPosition.Any())
            {
                return;
            }

            var argumentList = argumentsWithNewPosition[0].argumentReference.ArgumentListContext;
            var arguments = argumentList.argument();
            foreach (var (argumentReference, newIndex) in argumentsWithNewPosition)
            {
                if (argumentReference.ArgumentPosition == newIndex)
                {
                    continue;
                }

                var replacementArgument = argumentReference.Context.GetText();
                var contextToReplace = arguments[newIndex];

                if (contextToReplace.missingArgument() != null)
                {
                    //Missing members have are empty and thus stopIndex < startIndex, which is not legal for replace. 
                    rewriter.InsertBefore(contextToReplace.start.TokenIndex, replacementArgument);
                }
                else
                {
                    rewriter.Replace(contextToReplace, replacementArgument);
                }
            }
        }

        private static void AdjustSignatures(ReorderParametersModel model, IRewriteSession rewriteSession)
        {
            if (!model.Parameters.Any())
            {
                return;
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetDeclaration.QualifiedModuleName);
            var parameterList = model.Parameters[0].Declaration.Context.GetAncestor<VBAParser.ArgListContext>();
            var newIndices = NewIndicesOfParameterIndices(model);

            ReorderParameters(parameterList, newIndices, rewriter);
        }

        private static void ReorderParameters(VBAParser.ArgListContext parameterList, int[] newIndices, IModuleRewriter rewriter)
        {
            if (!newIndices.Any())
            {
                return;
            }

            var parameterContexts = parameterList.arg();
            for (var oldIndex = 0; oldIndex < newIndices.Length; oldIndex++)
            {
                var newIndex = newIndices[oldIndex];

                if (oldIndex == newIndex)
                {
                    continue;
                }

                var contextToReplace = parameterContexts[newIndex];
                var replacementParameter = parameterContexts[oldIndex].GetText();
                rewriter.Replace(contextToReplace, replacementParameter);
            }
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(declarationType)
                .FirstOrDefault(item => item.QualifiedModuleName.Equals(declaration.QualifiedModuleName)
                                        && item.IdentifierName == declaration.IdentifierName);
        }

        private static ReorderParametersModel ModelForNewTarget(ReorderParametersModel oldModel, Declaration newTarget)
        {
            var newModel = new ReorderParametersModel(newTarget);
            var newParameters = newModel.Parameters;

            var newReorderedParameters = oldModel.Parameters
                .Select(param => newParameters[param.Index])
                .ToList();
            if (newReorderedParameters.Count < newParameters.Count)
            {
                var additionalParameters = newParameters.Skip(newReorderedParameters.Count);
                newReorderedParameters.AddRange(additionalParameters);
            }

            newModel.Parameters = newReorderedParameters;
            return newModel;
        }
    }
}