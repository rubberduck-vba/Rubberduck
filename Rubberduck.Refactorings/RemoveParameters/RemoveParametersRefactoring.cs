using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.RemoveParameter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersRefactoring : InteractiveRefactoringBase<IRemoveParametersPresenter, RemoveParametersModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public RemoveParametersRefactoring(IDeclarationFinderProvider declarationFinderProvider, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            return _declarationFinderProvider.DeclarationFinder
                .AllUserDeclarations
                .FindTarget(targetSelection, ValidDeclarationTypes);
        }

        protected override RemoveParametersModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!ValidDeclarationTypes.Contains(target.DeclarationType) && target.DeclarationType != DeclarationType.Parameter)
            {
                throw new InvalidDeclarationTypeException(target);
            }

            var model = DerivedTarget(new RemoveParametersModel(target));

            return model;
        }

        private RemoveParametersModel DerivedTarget(RemoveParametersModel model)
        {
            var preliminaryModel = ResolvedInterfaceMemberTarget(model) 
                                   ?? ResolvedEventTarget(model) 
                                   ?? model;
            return ResolvedGetterTarget(preliminaryModel) ?? preliminaryModel;
        }

        private static RemoveParametersModel ResolvedInterfaceMemberTarget(RemoveParametersModel model)
        {
            var declaration = model.TargetDeclaration;
            if (!(declaration is ModuleBodyElementDeclaration member) || !member.IsInterfaceImplementation)
            {
                return null;
            }

            model.IsInterfaceMemberRefactoring = true;
            model.TargetDeclaration = member.InterfaceMemberImplemented;

            return model;
        }

        private RemoveParametersModel ResolvedEventTarget(RemoveParametersModel model)
        {
            foreach (var events in _declarationFinderProvider
                .DeclarationFinder
                .UserDeclarations(DeclarationType.Event))
            {
                if (_declarationFinderProvider.DeclarationFinder
                    .AllUserDeclarations
                    .FindHandlersForEvent(events)
                    .Any(reference => Equals(reference.Item2, model.TargetDeclaration)))
                {
                    model.IsEventRefactoring = true;
                    model.TargetDeclaration = events;
                    return model;
                }
            }
            return null;
        }

        private RemoveParametersModel ResolvedGetterTarget(RemoveParametersModel model)
        {
            var target = model.TargetDeclaration;
            if (target == null || !target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                return null;
            }

            if (target.DeclarationType == DeclarationType.PropertyGet)
            {
                model.IsPropertyRefactoringWithGetter = true;
                return model;
            }


            var getter = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.PropertyGet)
                .FirstOrDefault(item => item.Scope == target.Scope 
                                        && item.IdentifierName == target.IdentifierName);

            if (getter == null)
            {
                return null;
            }

            model.IsPropertyRefactoringWithGetter = true;
            model.TargetDeclaration = getter;

            return model;
        }

        protected override void RefactorImpl(RemoveParametersModel model)
        {
            RemoveParameters(model);
        }

        public void QuickFix(QualifiedSelection selection)
        {
            var targetDeclaration = FindTargetDeclaration(selection);
            var model = InitializeModel(targetDeclaration);
            
            var selectedParameters = model.Parameters.Where(p => selection.Selection.Contains(p.Declaration.QualifiedSelection.Selection)).ToList();

            if (selectedParameters.Count > 1)
            {
                throw new MultipleParametersSelectedException(selectedParameters);
            }

            var target = selectedParameters.SingleOrDefault(p => selection.Selection.Contains(p.Declaration.QualifiedSelection.Selection));

            if (target == null)
            {
                throw new NoParameterSelectedException();
            }

            model.RemoveParameters.Add(target);
            RemoveParameters(model);
        }

        private void RemoveParameters(RemoveParametersModel model)
        {
            if (model.TargetDeclaration == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();

            AdjustReferences(model, model.TargetDeclaration.References, model.TargetDeclaration, rewriteSession);
            AdjustSignatures(model, rewriteSession);

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private void AdjustReferences(RemoveParametersModel model, IEnumerable<IdentifierReference> references, Declaration method, IRewriteSession rewriteSession)
        {
            foreach (var reference in references.Where(item => item.Context != method.Context))
            {
                VBAParser.ArgumentListContext argumentList = null;
                var callStmt = reference.Context.GetAncestor<VBAParser.CallStmtContext>();
                if (callStmt != null)
                {
                    argumentList = CallStatement.GetArgumentList(callStmt);
                }

                if (argumentList == null)
                {
                    var indexExpression =
                        reference.Context.GetAncestor<VBAParser.IndexExprContext>();
                    if (indexExpression != null)
                    {
                        argumentList = indexExpression.GetChild<VBAParser.ArgumentListContext>();
                    }
                }

                if (argumentList == null)
                {
                    var whitespaceIndexExpression =
                        reference.Context.GetAncestor<VBAParser.WhitespaceIndexExprContext>();
                    if (whitespaceIndexExpression != null)
                    {
                        argumentList = whitespaceIndexExpression.GetChild<VBAParser.ArgumentListContext>();
                    }
                }

                if (argumentList == null)
                {
                    continue;
                }

                RemoveCallArguments(model, argumentList, reference.QualifiedModuleName, rewriteSession);
            }
        }

        private void RemoveCallArguments(RemoveParametersModel model, VBAParser.ArgumentListContext argList, QualifiedModuleName module, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(module);

            var usesNamedArguments = false;
            var args = argList.children.OfType<VBAParser.ArgumentContext>().ToList();
            for (var i = 0; i < model.Parameters.Count; i++)
            {
                // only remove params from RemoveParameters
                if (!model.RemoveParameters.Contains(model.Parameters[i]))
                {
                    continue;
                }
                
                if (model.Parameters[i].IsParamArray)
                {
                    //The following code works because it is neither allowed to use both named arguments
                    //and a ParamArray nor optional arguments and a ParamArray.
                    var index = i == 0 ? 0 : argList.children.IndexOf(args[i - 1]) + 1;
                    for (var j = index; j < argList.children.Count; j++)
                    {
                        rewriter.Remove(argList.children[j]);
                    }
                    break;
                }

                if (args.Count > i && (args[i].positionalArgument() != null || args[i].missingArgument() != null))
                {
                    rewriter.Remove(args[i]);
                }
                else
                {
                    usesNamedArguments = true;
                    var arg = args.Where(a => a.namedArgument() != null)
                                  .SingleOrDefault(a =>
                                        a.namedArgument().unrestrictedIdentifier().GetText() ==
                                        model.Parameters[i].Declaration.IdentifierName);

                    if (arg != null)
                    {
                        rewriter.Remove(arg);
                    }
                }
            }

            RemoveTrailingComma(model, rewriter, argList, usesNamedArguments);
        }

        private void AdjustSignatures(RemoveParametersModel model, IRewriteSession rewriteSession)
        {
            // if we are adjusting a property getter, check if we need to adjust the letter/setter too
            if (model.TargetDeclaration.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = GetLetterOrSetter(model.TargetDeclaration, DeclarationType.PropertySet);
                if (setter != null)
                {
                    RemoveSignatureParameters(model, setter, rewriteSession);
                    AdjustReferences(model, setter.References, setter, rewriteSession);
                }

                var letter = GetLetterOrSetter(model.TargetDeclaration, DeclarationType.PropertyLet);
                if (letter != null)
                {
                    RemoveSignatureParameters(model, letter, rewriteSession);
                    AdjustReferences(model, letter.References, letter, rewriteSession);
                }
            }

            RemoveSignatureParameters(model, model.TargetDeclaration, rewriteSession);

            var eventImplementations = _declarationFinderProvider.DeclarationFinder
                .AllUserDeclarations
                .Where(item => item.IsWithEvents && item.AsTypeName == model.TargetDeclaration.ComponentName)
                .SelectMany(withEvents => _declarationFinderProvider.DeclarationFinder
                    .AllUserDeclarations.FindEventProcedures(withEvents));

            foreach (var eventImplementation in eventImplementations)
            {
                AdjustReferences(model, eventImplementation.References, eventImplementation, rewriteSession);
                RemoveSignatureParameters(model, eventImplementation, rewriteSession);
            }

            var interfaceImplementations = _declarationFinderProvider.DeclarationFinder
                .FindAllInterfaceImplementingMembers()
                .Where(item => item.ProjectId == model.TargetDeclaration.ProjectId
                               && item.IdentifierName == $"{model.TargetDeclaration.ComponentName}_{model.TargetDeclaration.IdentifierName}");

            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustReferences(model, interfaceImplentation.References, interfaceImplentation, rewriteSession);
                RemoveSignatureParameters(model, interfaceImplentation, rewriteSession);
            }
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(declarationType)
                .FirstOrDefault(item => item.QualifiedModuleName.Equals(declaration.QualifiedModuleName)
                && item.IdentifierName == declaration.IdentifierName);
        }

        private void RemoveSignatureParameters(RemoveParametersModel model, Declaration target, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            var parameters = ((IParameterizedDeclaration)target).Parameters.OrderBy(o => o.Selection).ToList();

            foreach (var index in model.RemoveParameters.Select(rem => model.Parameters.IndexOf(rem)))
            {
                rewriter.Remove(parameters[index]);
            }

            RemoveTrailingComma(model, rewriter);
        }

        //Issue 4319.  If there are 3 or more arguments and the user elects to remove 2 or more of
        //the last arguments, then we need to specifically remove the trailing comma from
        //the last 'kept' argument.
        private void RemoveTrailingComma(RemoveParametersModel model, IModuleRewriter rewriter, VBAParser.ArgumentListContext argList = null, bool usesNamedParams = false)
        {
            var commaLocator = RetrieveTrailingCommaInfo(model.RemoveParameters, model.Parameters);
            if (!commaLocator.RequiresTrailingCommaRemoval)
            {
                return;
            }

            var tokenStart = 0;
            var tokenStop = 0;

            if (argList is null)
            {
                //Handle Signatures only
                tokenStart = commaLocator.LastRetainedArg.Param.Declaration.Context.Stop.TokenIndex + 1;
                tokenStop = commaLocator.FirstOfRemovedArgSeries.Param.Declaration.Context.Start.TokenIndex - 1;
                rewriter.RemoveRange(tokenStart, tokenStop);
                return;
            }


            //Handles References
            var args = argList.children.OfType<VBAParser.ArgumentContext>().ToList();

            if (usesNamedParams)
            {
                var lastKeptArg = args.Where(a => a.namedArgument() != null)
                    .SingleOrDefault(a => a.namedArgument().unrestrictedIdentifier().GetText() ==
                                            commaLocator.LastRetainedArg.Identifier);

                var firstOfRemovedArgSeries = args.Where(a => a.namedArgument() != null)
                    .SingleOrDefault(a => a.namedArgument().unrestrictedIdentifier().GetText() ==
                                            commaLocator.FirstOfRemovedArgSeries.Identifier);

                tokenStart = lastKeptArg.Stop.TokenIndex + 1;
                tokenStop = firstOfRemovedArgSeries.Start.TokenIndex - 1;
                rewriter.RemoveRange(tokenStart, tokenStop);
                return;
            }
            tokenStart = args[commaLocator.LastRetainedArg.Index].Stop.TokenIndex + 1;
            tokenStop = args[commaLocator.FirstOfRemovedArgSeries.Index].Start.TokenIndex - 1;
            rewriter.RemoveRange(tokenStart, tokenStop);
        }

        private static CommaLocator RetrieveTrailingCommaInfo(List<Parameter> toRemove, List<Parameter> allParams)
        {
            if (toRemove.Count == allParams.Count || allParams.Count < 3)
            {
                return new CommaLocator();
            }

            var reversedAllParams = allParams.OrderByDescending(tr => tr.Declaration.Selection).ToList();
            var rangeRemoval = new List<Parameter>();
            for (var idx = 0; idx < reversedAllParams.Count(); idx++)
            {
                if (toRemove.Contains(reversedAllParams.ElementAt(idx)))
                {
                    rangeRemoval.Add(reversedAllParams.ElementAt(idx));
                    continue;
                }

                if (rangeRemoval.Count >= 2)
                {
                    var startIndex = allParams.FindIndex(par => par == reversedAllParams.ElementAt(idx));
                    var stopIndex = allParams.FindIndex(par => par == rangeRemoval.First());

                    return new CommaLocator()
                    {
                        RequiresTrailingCommaRemoval = true,
                        LastRetainedArg = new CommaBoundary()
                        {
                            Param = reversedAllParams.ElementAt(idx),
                            Index = startIndex,
                        },
                        FirstOfRemovedArgSeries = new CommaBoundary()
                        {
                            Param = rangeRemoval.First(),
                            Index = stopIndex,
                        }
                    };
                }
                break;
            }
            return new CommaLocator();
        }

        private struct CommaLocator
        {
            public bool RequiresTrailingCommaRemoval;
            public CommaBoundary LastRetainedArg;
            public CommaBoundary FirstOfRemovedArgSeries;
        }

        private struct CommaBoundary
        {
            public Parameter Param;
            public int Index;
            public string Identifier => Param.Declaration.IdentifierName;
        }

        public static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };
    }
}
