﻿using System;
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

            var parameters = ((IParameterizedDeclaration) target).Parameters.OrderBy(o => o.Selection).ToList();
            
            foreach (var index in _model.RemoveParameters.Select(rem => _model.Parameters.IndexOf(rem)))
            {
                rewriter.Remove(parameters[index]);
            }

            _rewriters.Add(rewriter);
        }
    }
}
