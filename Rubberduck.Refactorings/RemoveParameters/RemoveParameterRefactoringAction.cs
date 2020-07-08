using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParameterRefactoringAction : RefactoringActionBase<RemoveParametersModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public RemoveParameterRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager) 
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        protected override void Refactor(RemoveParametersModel model, IRewriteSession rewriteSession)
        {
            AdjustReferences(model, rewriteSession);
            AdjustSignatures(model, rewriteSession);

            // if we are adjusting a property getter, check if we need to adjust the letter/setter too
            if (model.TargetDeclaration.DeclarationType == DeclarationType.PropertyGet)
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

            foreach (var eventImplementation in eventImplementations)
            {
                var eventImplementationModel = ModelForNewTarget(model, eventImplementation);
                AdjustReferences(eventImplementationModel, rewriteSession);
                AdjustSignatures(eventImplementationModel, rewriteSession);
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

        private static void AdjustReferences(RemoveParametersModel model, IRewriteSession rewriteSession)
        {
            var parametersToRemove = model.RemoveParameters
                .Select(parameter => parameter.Declaration)
                .ToList();
            var argumentReferences = ArgumentReferencesByLocation(parametersToRemove);

            foreach (var (module, moduleArgumentReferences) in argumentReferences)
            {
                AdjustReferences(module, moduleArgumentReferences, rewriteSession);
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
            QualifiedModuleName module,
            Dictionary<Selection, List<ArgumentReference>> argumentReferences,
            IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(module);
            foreach (var (argumentListSelection, sameArgumentListReferences) in argumentReferences)
            {
                //This happens for (with) dictionary access expressions only.
                if (argumentListSelection.Equals(Selection.Empty))
                {
                    foreach (var dictionaryAccessArgument in sameArgumentListReferences)
                    {
                        ReplaceDictionaryAccess(dictionaryAccessArgument, rewriter);
                    }

                    continue;
                }

                AdjustReferences(sameArgumentListReferences, rewriter);
            }
        }

        private static void ReplaceDictionaryAccess(ArgumentReference dictionaryAccessArgument, IModuleRewriter rewriter)
        {
            //TODO: Deal with WithDictionaryAccessExprContexts.
            //This should best be handled by extracting a refactoring out of the ExpandBangNotationQuickFix and
            //using it here to expand the dictionary access for both kinds of dictionary access expression.
            var dictionaryAccess = dictionaryAccessArgument?.Context?.Parent as VBAParser.DictionaryAccessExprContext;

            if (dictionaryAccess == null)
            {
                return;
            }

            var startTokenIndex = dictionaryAccess.dictionaryAccess().start.TokenIndex;
            var stopTokenIndex = dictionaryAccess.unrestrictedIdentifier().stop.TokenIndex;
            const string replacementString = "()";
            rewriter.Replace(new Interval(startTokenIndex, stopTokenIndex), replacementString);
        }

        private static void AdjustReferences(IReadOnlyCollection<ArgumentReference> argumentReferences, IModuleRewriter rewriter)
        {
            if (!argumentReferences.Any())
            {
                return;
            }

            var argumentIndicesToRemove = argumentReferences.Select(argumentReference => argumentReference.ArgumentPosition);
            var argumentIndexRangesToRemove = IndexRanges(argumentIndicesToRemove);
            var argumentList = argumentReferences.First().ArgumentListContext;

            var adjustedArgumentIndexRangesToRemove = WithTrailingMissingArguments(argumentIndexRangesToRemove, argumentList);


            foreach (var (startIndex, stopIndex) in adjustedArgumentIndexRangesToRemove)
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

        private static IEnumerable<(int startIndex, int stopIndex)> WithTrailingMissingArguments(
            IEnumerable<(int startIndex, int stopIndex)> argumentRanges,
            VBAParser.ArgumentListContext argumentList)
        {
            var arguments = argumentList.argument();
            var numberOfArguments = arguments.Length;

            var argumentRangesInDescendingOrder = argumentRanges.OrderByDescending(range => range.stopIndex).ToList();
            if (argumentRangesInDescendingOrder[0].stopIndex != numberOfArguments - 1)
            {
                return argumentRangesInDescendingOrder;
            }

            var currentRangeIndex = 0;
            var currentStartIndex = argumentRangesInDescendingOrder[0].startIndex;
            while (currentStartIndex > 0)
            {
                if (currentRangeIndex + 1 < argumentRangesInDescendingOrder.Count
                    && currentStartIndex - 1 == argumentRangesInDescendingOrder[currentRangeIndex + 1].stopIndex)
                {
                    currentRangeIndex++;
                    currentStartIndex = argumentRangesInDescendingOrder[currentRangeIndex].startIndex;
                }
                else if (arguments[currentStartIndex - 1]?.missingArgument() != null)
                {
                    currentStartIndex--;
                }
                else
                {
                    break;
                }
            }

            var newRanges = new List<(int startIndex, int stopIndex)> { (currentStartIndex, numberOfArguments - 1) };
            newRanges.AddRange(argumentRangesInDescendingOrder.Skip(currentRangeIndex + 1));
            return newRanges;
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

        private static void AdjustSignatures(RemoveParametersModel model, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetDeclaration.QualifiedModuleName);

            var parameterIndicesToRemove = model.RemoveParameters
                .Select(param => model.Parameters.IndexOf(param));
            var parameterRangesToRemove = IndexRanges(parameterIndicesToRemove);

            var argList = model.Parameters.First().Declaration.Context.GetAncestor<VBAParser.ArgListContext>();

            foreach (var (startIndex, stopIndex) in parameterRangesToRemove)
            {
                RemoveParameterRange(startIndex, stopIndex, argList, rewriter);
            }
        }

        private static void RemoveParameterRange(
            int startArgumentIndex,
            int stopArgumentIndex,
            VBAParser.ArgListContext argList,
            IModuleRewriter rewriter)
        {
            var (startTokenIndex, stopTokenIndex) = TokenIndexRange(startArgumentIndex, stopArgumentIndex, argList.arg());
            rewriter.RemoveRange(startTokenIndex, stopTokenIndex);
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(declarationType)
                .FirstOrDefault(item => item.QualifiedModuleName.Equals(declaration.QualifiedModuleName)
                                        && item.IdentifierName == declaration.IdentifierName);
        }

        private static RemoveParametersModel ModelForNewTarget(RemoveParametersModel oldModel, Declaration newTarget)
        {
            var newModel = new RemoveParametersModel(newTarget);
            var toRemoveIndices = oldModel.RemoveParameters.Select(param => oldModel.Parameters.IndexOf(param));
            var newToRemoveParams = toRemoveIndices
                .Select(index => newModel.Parameters[index])
                .ToList();
            newModel.RemoveParameters = newToRemoveParams;
            return newModel;
        }
    }
}