using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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

        public override void Refactor(QualifiedSelection target)
        {
            Refactor(InitializeModel(target));
        }

        private RemoveParametersModel InitializeModel(QualifiedSelection targetSelection)
        {
            return new RemoveParametersModel(_declarationFinderProvider, targetSelection);
        }

        public override void Refactor(Declaration target)
        {
            Refactor(InitializeModel(target));
        }

        private RemoveParametersModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                return null;
            }

            if (!RemoveParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType) && target.DeclarationType != DeclarationType.Parameter)
            {
                return null;
            }

            return InitializeModel(target.QualifiedSelection);
        }

        protected override void RefactorImpl(IRemoveParametersPresenter presenter)
        {
            RemoveParameters();
        }

        public void QuickFix(QualifiedSelection selection)
        {
            Model = InitializeModel(selection);
            
            var target = Model.Parameters.SingleOrDefault(p => selection.Selection.Contains(p.Declaration.QualifiedSelection.Selection));
            Debug.Assert(target != null, "Target was not found");
            
            if (target != null)
            {
                Model.RemoveParameters.Add(target);
            }
            else
            {
                return;
            }
            RemoveParameters();
        }

        private void RemoveParameters()
        {
            if (Model.TargetDeclaration == null)
            {
                throw new NullReferenceException("Parameter is null");
            }

            var rewritingSession = RewritingManager.CheckOutCodePaneSession();

            AdjustReferences(Model.TargetDeclaration.References, Model.TargetDeclaration, rewritingSession);
            AdjustSignatures(rewritingSession);

            rewritingSession.TryRewrite();
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references, Declaration method, IRewriteSession rewriteSession)
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

                RemoveCallArguments(argumentList, reference.QualifiedModuleName, rewriteSession);
            }
        }

        private void RemoveCallArguments(VBAParser.ArgumentListContext argList, QualifiedModuleName module, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(module);

            var usesNamedArguments = false;
            var args = argList.children.OfType<VBAParser.ArgumentContext>().ToList();
            for (var i = 0; i < Model.Parameters.Count; i++)
            {
                // only remove params from RemoveParameters
                if (!Model.RemoveParameters.Contains(Model.Parameters[i]))
                {
                    continue;
                }
                
                if (Model.Parameters[i].IsParamArray)
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
                                        Model.Parameters[i].Declaration.IdentifierName);

                    if (arg != null)
                    {
                        rewriter.Remove(arg);
                    }
                }
            }

            RemoveTrailingComma(rewriter, argList, usesNamedArguments);
        }

        private void AdjustSignatures(IRewriteSession rewriteSession)
        {
            // if we are adjusting a property getter, check if we need to adjust the letter/setter too
            if (Model.TargetDeclaration.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = GetLetterOrSetter(Model.TargetDeclaration, DeclarationType.PropertySet);
                if (setter != null)
                {
                    RemoveSignatureParameters(setter, rewriteSession);
                    AdjustReferences(setter.References, setter, rewriteSession);
                }

                var letter = GetLetterOrSetter(Model.TargetDeclaration, DeclarationType.PropertyLet);
                if (letter != null)
                {
                    RemoveSignatureParameters(letter, rewriteSession);
                    AdjustReferences(letter.References, letter, rewriteSession);
                }
            }

            RemoveSignatureParameters(Model.TargetDeclaration, rewriteSession);

            var eventImplementations = Model.Declarations
                .Where(item => item.IsWithEvents && item.AsTypeName == Model.TargetDeclaration.ComponentName)
                .SelectMany(withEvents => Model.Declarations.FindEventProcedures(withEvents));

            foreach (var eventImplementation in eventImplementations)
            {
                AdjustReferences(eventImplementation.References, eventImplementation, rewriteSession);
                RemoveSignatureParameters(eventImplementation, rewriteSession);
            }

            var interfaceImplementations = _declarationFinderProvider.DeclarationFinder
                .FindAllInterfaceImplementingMembers()
                .Where(item => item.ProjectId == Model.TargetDeclaration.ProjectId 
                               && item.IdentifierName == $"{Model.TargetDeclaration.ComponentName}_{Model.TargetDeclaration.IdentifierName}");

            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustReferences(interfaceImplentation.References, interfaceImplentation, rewriteSession);
                RemoveSignatureParameters(interfaceImplentation, rewriteSession);
            }
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return Model.Declarations.FirstOrDefault(item => item.QualifiedModuleName.Equals(declaration.QualifiedModuleName) 
                && item.IdentifierName == declaration.IdentifierName 
                && item.DeclarationType == declarationType);
        }

        private void RemoveSignatureParameters(Declaration target, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);

            var parameters = ((IParameterizedDeclaration) target).Parameters.OrderBy(o => o.Selection).ToList();
            
            foreach (var index in Model.RemoveParameters.Select(rem => Model.Parameters.IndexOf(rem)))
            {
                rewriter.Remove(parameters[index]);
            }

            RemoveTrailingComma(rewriter);
        }

        //Issue 4319.  If there are 3 or more arguments and the user elects to remove 2 or more of
        //the last arguments, then we need to specifically remove the trailing comma from
        //the last 'kept' argument.
        private void RemoveTrailingComma(IModuleRewriter rewriter, VBAParser.ArgumentListContext argList = null, bool usesNamedParams = false)
        {
            var commaLocator = RetrieveTrailingCommaInfo(Model.RemoveParameters, Model.Parameters);
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

        private CommaLocator RetrieveTrailingCommaInfo(List<Parameter> toRemove, List<Parameter> allParams)
        {
            if (toRemove.Count == allParams.Count || allParams.Count < 3)
            {
                return new CommaLocator();
            }

            var reversedAllParams = allParams.OrderByDescending(tr => tr.Declaration.Selection);
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
    }
}
