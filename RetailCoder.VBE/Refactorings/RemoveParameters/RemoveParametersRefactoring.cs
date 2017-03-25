using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
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
            if (_model == null || !_model.Parameters.Any(item => item.IsRemoved))
            {
                return;
            }

            QualifiedSelection? oldSelection = null;
            var pane = _vbe.ActiveCodePane;
            var module = pane.CodeModule;
            if (!module.IsWrappingNullReference)
            {
                oldSelection = module.GetQualifiedSelection();
            }

            RemoveParameters();

            if (oldSelection.HasValue)
            {
                pane.Selection = oldSelection.Value.Selection;
            }

            _model.State.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            var pane = _vbe.ActiveCodePane;
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

            var pane = _vbe.ActiveCodePane;
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

            target.IsRemoved = true;
            RemoveParameters();
        }

        private void RemoveParameters()
        {
            if (_model.TargetDeclaration == null) { throw new NullReferenceException("Parameter is null"); }

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
                var module = reference.QualifiedModuleName.Component.CodeModule;
                VBAParser.ArgumentListContext argumentList = null;
                var callStmt = ParserRuleContextHelper.GetParent<VBAParser.CallStmtContext>(reference.Context);
                if (callStmt != null)
                {
                    argumentList = CallStatement.GetArgumentList(callStmt);
                }

                if (argumentList == null)
                {
                    var indexExpression =
                        ParserRuleContextHelper.GetParent<VBAParser.IndexExprContext>(reference.Context);
                    if (indexExpression != null)
                    {
                        argumentList = ParserRuleContextHelper.GetChild<VBAParser.ArgumentListContext>(indexExpression);
                    }
                }

                if (argumentList == null)
                {
                    continue;
                }

                RemoveCallArguments(argumentList, module);
            }
        }

        private void RemoveCallArguments(VBAParser.ArgumentListContext argList, ICodeModule module)
        {
            var rewriter = _model.State.GetRewriter(module.Parent);

            var args = argList.children.OfType<VBAParser.ArgumentContext>().ToList();
            for (var i = 0; i < _model.Parameters.Count; i++)
            {
                if (!_model.Parameters[i].IsRemoved)
                {
                    continue;
                }
                
                if (_model.Parameters[i].IsParamArray)
                {
                    var index = i == 0 ? 0 : argList.children.IndexOf(args[i - 1]) + 1;
                    for (var j = index; j < argList.children.Count; j++)
                    {
                        rewriter.Remove((dynamic)argList.children[j]);
                    }
                    break;
                }

                if (args.Count > i && args[i].positionalArgument() != null)
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

            var interfaceImplementations = _model.Declarations.FindInterfaceImplementationMembers().Where(item =>
                item.ProjectId == _model.TargetDeclaration.ProjectId
                &&
                item.IdentifierName ==
                _model.TargetDeclaration.ComponentName + "_" + _model.TargetDeclaration.IdentifierName);

            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustReferences(interfaceImplentation.References, interfaceImplentation);
                RemoveSignatureParameters(interfaceImplentation);
            }
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return _model.Declarations.FirstOrDefault(item => item.Scope == declaration.Scope 
                && item.IdentifierName == declaration.IdentifierName 
                && item.DeclarationType == declarationType);
        }

        private void RemoveSignatureParameters(Declaration target)
        {
            var rewriter = _model.State.GetRewriter(target);

            var parameters = ((IParameterizedDeclaration) target).Parameters.OrderBy(o => o.Selection).ToList();
            
            for (var i = 0; i < _model.Parameters.Count; i++)
            {
                if (_model.Parameters[i].IsRemoved)
                {
                    rewriter.Remove(parameters[i]);
                }
            }

            _rewriters.Add(rewriter);
        }
    }
}
