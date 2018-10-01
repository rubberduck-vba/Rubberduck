using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringPresenterFactory<IRemoveParametersPresenter> _factory;
        private RemoveParametersModel _model;
        private readonly HashSet<IModuleRewriter> _rewriters = new HashSet<IModuleRewriter>();

        public RemoveParametersRefactoring(IVBE vbe, IRefactoringPresenterFactory<IRemoveParametersPresenter> factory)
        {
            _vbe = vbe;
            _factory = factory;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                return;
            }

            _model = presenter.Show();
            if (_model == null || !_model.Parameters.Any())
            {
                return;
            }

            using (var pane = _vbe.ActiveCodePane)
            {
                var oldSelection = pane.GetQualifiedSelection();

                RemoveParameters();

                if (oldSelection.HasValue && !pane.IsWrappingNullReference)
                {
                    pane.Selection = oldSelection.Value.Selection;
                }
            }

            _model.State.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }
                pane.Selection = target.Selection;
                Refactor();
            }
        }

        public void Refactor(Declaration target)
        {
            if (!RemoveParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType) && target.DeclarationType != DeclarationType.Parameter)
            {
                throw new ArgumentException("Invalid declaration type");
            }

            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }
                pane.Selection = target.QualifiedSelection.Selection;
                Refactor();
            }
        }

        public void QuickFix(RubberduckParserState state, QualifiedSelection selection)
        {
            _model = new RemoveParametersModel(state, selection, new MessageBox());

            var target = _model.Parameters.SingleOrDefault(p => selection.Selection.Contains(p.Declaration.QualifiedSelection.Selection));
            Debug.Assert(target != null, "Target was not found");

            if (target != null)
            {
                _model.RemoveParameters.Add(target);
            } else
            {
                return;
            }
            RemoveParameters();
        }

        private void RemoveParameters()
        {
            if (_model.TargetDeclaration == null)
            {
                throw new NullReferenceException("Parameter is null");
            }

            AdjustReferences(_model.TargetDeclaration.References, _model.TargetDeclaration);
            AdjustSignatures();

            foreach (var rewriter in _rewriters)
            {
                rewriter.Rewrite();
            }
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references, Declaration method)
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

                RemoveCallArguments(argumentList, reference.QualifiedModuleName);
            }
        }

        private void RemoveCallArguments(VBAParser.ArgumentListContext argList, QualifiedModuleName module)
        {
            var rewriter = _model.State.GetRewriter(module);
            _rewriters.Add(rewriter);

            var usesNamedArguments = false;
            var args = argList.children.OfType<VBAParser.ArgumentContext>().ToList();
            for (var i = 0; i < _model.Parameters.Count; i++)
            {
                // only remove params from RemoveParameters
                if (!_model.RemoveParameters.Contains(_model.Parameters[i]))
                {
                    continue;
                }

                if (_model.Parameters[i].IsParamArray)
                {
                    //The following code works because it is neither allowed to use both named arguments
                    //and a ParamArray nor optional arguments and a ParamArray.
                    var index = i == 0 ? 0 : argList.children.IndexOf(args[i - 1]) + 1;
                    for (var j = index; j < argList.children.Count; j++)
                    {
                        rewriter.Remove((dynamic)argList.children[j]);
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
                                        _model.Parameters[i].Declaration.IdentifierName);

                    if (arg != null)
                    {
                        rewriter.Remove(arg);
                    }
                }
            }

            RemoveTrailingComma(rewriter, argList, usesNamedArguments);
        }

        private void AdjustSignatures()
        {
            // if we are adjusting a property getter, check if we need to adjust the letter/setter too
            if (_model.TargetDeclaration.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = GetLetterOrSetter(_model.TargetDeclaration, DeclarationType.PropertySet);
                if (setter != null)
                {
                    RemoveSignatureParameters(setter);
                    AdjustReferences(setter.References, setter);
                }

                var letter = GetLetterOrSetter(_model.TargetDeclaration, DeclarationType.PropertyLet);
                if (letter != null)
                {
                    RemoveSignatureParameters(letter);
                    AdjustReferences(letter.References, letter);
                }
            }

            RemoveSignatureParameters(_model.TargetDeclaration);

            var eventImplementations = _model.Declarations
                .Where(item => item.IsWithEvents && item.AsTypeName == _model.TargetDeclaration.ComponentName)
                .SelectMany(withEvents => _model.Declarations.FindEventProcedures(withEvents));

            foreach (var eventImplementation in eventImplementations)
            {
                AdjustReferences(eventImplementation.References, eventImplementation);
                RemoveSignatureParameters(eventImplementation);
            }

            var interfaceImplementations = _model.State.DeclarationFinder.FindAllInterfaceImplementingMembers().Where(item =>
                item.ProjectId == _model.TargetDeclaration.ProjectId
                &&
                item.IdentifierName == $"{_model.TargetDeclaration.ComponentName}_{_model.TargetDeclaration.IdentifierName}");

            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustReferences(interfaceImplentation.References, interfaceImplentation);
                RemoveSignatureParameters(interfaceImplentation);
            }
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return _model.Declarations.FirstOrDefault(item => item.QualifiedModuleName.Equals(declaration.QualifiedModuleName)
                && item.IdentifierName == declaration.IdentifierName
                && item.DeclarationType == declarationType);
        }

        private void RemoveSignatureParameters(Declaration target)
        {
            var rewriter = _model.State.GetRewriter(target);

            var parameters = ((IParameterizedDeclaration)target).Parameters.OrderBy(o => o.Selection).ToList();

            foreach (var index in _model.RemoveParameters.Select(rem => _model.Parameters.IndexOf(rem)))
            {
                rewriter.Remove(parameters[index]);
            }

            RemoveTrailingComma(rewriter);
            _rewriters.Add(rewriter);
        }

        //Issue 4319.  If there are 3 or more arguments and the user elects to remove 2 or more of
        //the last arguments, then we need to specifically remove the trailing comma from
        //the last 'kept' argument.
        private void RemoveTrailingComma(IModuleRewriter rewriter, VBAParser.ArgumentListContext argList = null, bool usesNamedParams = false)
        {
            var commaLocator = RetrieveTrailingCommaInfo(_model.RemoveParameters, _model.Parameters);
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
