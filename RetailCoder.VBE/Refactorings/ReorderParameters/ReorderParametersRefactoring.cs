using Antlr4.Runtime.Misc;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringPresenterFactory<IReorderParametersPresenter> _factory;
        private ReorderParametersModel _model;
        private readonly IMessageBox _messageBox;

        public ReorderParametersRefactoring(IVBE vbe, IRefactoringPresenterFactory<IReorderParametersPresenter> factory, IMessageBox messageBox)
        {
            _vbe = vbe;
            _factory = factory;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                return;
            }

            _model = presenter.Show();
            if (_model == null || !_model.Parameters.Where((param, index) => param.Index != index).Any() || !IsValidParamOrder())
            {
                return;
            }

            var pane = _vbe.ActiveCodePane;
            if (!pane.IsWrappingNullReference)
            {
                QualifiedSelection? oldSelection;
                var module = pane.CodeModule;
                {
                    oldSelection = module.GetQualifiedSelection();
                }

                AdjustReferences(_model.TargetDeclaration.References);
                AdjustSignatures();

                if (oldSelection.HasValue)
                {
                    pane.Selection = oldSelection.Value.Selection;
                }
            }

            _model.State.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            var pane = _vbe.ActiveCodePane;
            {
                pane.Selection = target.Selection;
                Refactor();
            }
        }

        public void Refactor(Declaration target)
        {
            if (!ReorderParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                throw new ArgumentException("Invalid declaration type");
            }

            var pane = _vbe.ActiveCodePane;
            {
                pane.Selection = target.QualifiedSelection.Selection;
                Refactor();
            }
        }

        private bool IsValidParamOrder()
        {
            var indexOfFirstOptionalParam = _model.Parameters.FindIndex(param => param.IsOptional);
            if (indexOfFirstOptionalParam >= 0)
            {
                for (var index = indexOfFirstOptionalParam + 1; index < _model.Parameters.Count; index++)
                {
                    if (!_model.Parameters.ElementAt(index).IsOptional)
                    {
                        _messageBox.Show(RubberduckUI.ReorderPresenter_OptionalParametersMustBeLastError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return false;
                    }
                }
            }

            var indexOfParamArray = _model.Parameters.FindIndex(param => param.IsParamArray);
            if (indexOfParamArray >= 0 && indexOfParamArray != _model.Parameters.Count - 1)
            {
                _messageBox.Show(RubberduckUI.ReorderPresenter_ParamArrayError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            return true;
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references)
        {
            foreach (var reference in references.Where(item => item.Context != _model.TargetDeclaration.Context))
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
                    var indexExpression = ParserRuleContextHelper.GetParent<VBAParser.IndexExprContext>(reference.Context);
                    if (indexExpression != null)
                    {
                        argumentList = ParserRuleContextHelper.GetChild<VBAParser.ArgumentListContext>(indexExpression);
                    }
                }

                if (argumentList == null) { continue; }
                RewriteCall(argumentList, module);
            }
        }

        private void RewriteCall(VBAParser.ArgumentListContext paramList, ICodeModule module)
        {
            var argValues = new List<string>();
            if (paramList.positionalOrNamedArgumentList().positionalArgumentOrMissing() != null)
            {
                argValues.AddRange(paramList.positionalOrNamedArgumentList().positionalArgumentOrMissing().Select(p =>
                {
                    if (p is VBAParser.SpecifiedPositionalArgumentContext)
                    {
                        return ((VBAParser.SpecifiedPositionalArgumentContext)p).positionalArgument().GetText();
                    }

                    return string.Empty;
                }).ToList());
            }
            if (paramList.positionalOrNamedArgumentList().namedArgumentList() != null)
            {
                argValues.AddRange(paramList.positionalOrNamedArgumentList().namedArgumentList().namedArgument().Select(p => p.GetText()).ToList());
            }
            if (paramList.positionalOrNamedArgumentList().requiredPositionalArgument() != null)
            {
                argValues.Add(paramList.positionalOrNamedArgumentList().requiredPositionalArgument().GetText());
            }

            var lineCount = paramList.Stop.Line - paramList.Start.Line + 1; // adjust for total line count

            var newContent = module.GetLines(paramList.Start.Line, lineCount);
            newContent = newContent.Remove(paramList.Start.Column, paramList.GetText().Length);

            var reorderedArgValues = new List<string>();
            foreach (var param in _model.Parameters)
            {
                var argAtIndex = argValues.ElementAtOrDefault(param.Index);
                if (argAtIndex != null)
                {
                    reorderedArgValues.Add(argAtIndex);
                }
            }

            // property let/set and paramarrays
            for (var index = reorderedArgValues.Count; index < argValues.Count; index++)
            {
                reorderedArgValues.Add(argValues[index]);
            }

            newContent = newContent.Insert(paramList.Start.Column, string.Join(", ", reorderedArgValues));

            module.ReplaceLine(paramList.Start.Line, newContent.Replace(" _" + Environment.NewLine, string.Empty));
            module.DeleteLines(paramList.Start.Line + 1, lineCount - 1);
        }

        private void AdjustSignatures()
        {
            var proc = (dynamic)_model.TargetDeclaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = _model.TargetDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;

            // if we are reordering a property getter, check if we need to reorder a letter/setter too
            if (_model.TargetDeclaration.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = _model.Declarations.FirstOrDefault(item => item.ParentScope == _model.TargetDeclaration.ParentScope &&
                                              item.IdentifierName == _model.TargetDeclaration.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertySet);

                if (setter != null)
                {
                    AdjustSignatures(setter);
                    AdjustReferences(setter.References);
                }

                var letter = _model.Declarations.FirstOrDefault(item => item.ParentScope == _model.TargetDeclaration.ParentScope &&
                              item.IdentifierName == _model.TargetDeclaration.IdentifierName &&
                              item.DeclarationType == DeclarationType.PropertyLet);

                if (letter != null)
                {
                    AdjustSignatures(letter);
                    AdjustReferences(letter.References);
                }
            }

            RewriteSignature(_model.TargetDeclaration, paramList, module);

            foreach (var withEvents in _model.Declarations.Where(item => item.IsWithEvents && item.AsTypeName == _model.TargetDeclaration.ComponentName))
            {
                foreach (var reference in _model.Declarations.FindEventProcedures(withEvents))
                {
                    AdjustReferences(reference.References);
                    AdjustSignatures(reference);
                }
            }

            var interfaceImplementations = _model.Declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.ProjectId == _model.TargetDeclaration.ProjectId &&
                                                               item.IdentifierName == _model.TargetDeclaration.ComponentName + "_" + _model.TargetDeclaration.IdentifierName);
            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustReferences(interfaceImplentation.References);
                AdjustSignatures(interfaceImplentation);
            }
        }

        private void AdjustSignatures(Declaration declaration)
        {
            var proc = (dynamic)declaration.Context.Parent;
            var module = declaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            {
                VBAParser.ArgListContext paramList;

                if (declaration.DeclarationType == DeclarationType.PropertySet || declaration.DeclarationType == DeclarationType.PropertyLet)
                {
                    paramList = (VBAParser.ArgListContext)proc.children[0].argList();
                }
                else
                {
                    paramList = (VBAParser.ArgListContext)proc.subStmt().argList();
                }

                RewriteSignature(declaration, paramList, module);
            }
        }

        private void RewriteSignature(Declaration target, VBAParser.ArgListContext paramList, ICodeModule module)
        {
            var parameters = paramList.arg().Select((s, i) => new {Index = i, Text = s.GetText()}).ToList();

            var reorderedParams = new List<string>();
            foreach (var param in _model.Parameters)
            {
                var parameterAtIndex = parameters.SingleOrDefault(s => s.Index == param.Index);
                if (parameterAtIndex != null)
                {
                    reorderedParams.Add(parameterAtIndex.Text);
                }
            }

            // property let/set and paramarrays
            for (var index = reorderedParams.Count; index < parameters.Count; index++)
            {
                reorderedParams.Add(parameters[index].Text);
            }

            var signature = GetOldSignature(target);
            signature = signature.Remove(signature.IndexOf('('));

            var asTypeText = target.AsTypeContext == null ? string.Empty : " " + target.AsTypeContext.GetText();
            signature += '(' + string.Join(", ", reorderedParams) + ")" + (asTypeText == " " ? string.Empty : asTypeText);

            var lineCount = paramList.GetSelection().LineCount;
            module.ReplaceLine(paramList.Start.Line, signature.Replace(" _" + Environment.NewLine, string.Empty));
            module.DeleteLines(paramList.Start.Line + 1, lineCount - 1);
        }

        private string GetOldSignature(Declaration target)
        {
            var rewriter = _model.State.GetRewriter(target.QualifiedName.QualifiedModuleName.Component);

            var context = target.Context;
            var firstTokenIndex = context.Start.TokenIndex;
            var lastTokenIndex = -1; // will blow up if this code runs for any context other than below

            var subStmtContext = context as VBAParser.SubStmtContext;
            if (subStmtContext != null)
            {
                lastTokenIndex = subStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            var functionStmtContext = context as VBAParser.FunctionStmtContext;
            if (functionStmtContext != null)
            {
                lastTokenIndex = functionStmtContext.asTypeClause() != null
                    ? functionStmtContext.asTypeClause().Stop.TokenIndex
                    : functionStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            var propertyGetStmtContext = context as VBAParser.PropertyGetStmtContext;
            if (propertyGetStmtContext != null)
            {
                lastTokenIndex = propertyGetStmtContext.asTypeClause() != null
                    ? propertyGetStmtContext.asTypeClause().Stop.TokenIndex
                    : propertyGetStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            var propertyLetStmtContext = context as VBAParser.PropertyLetStmtContext;
            if (propertyLetStmtContext != null)
            {
                lastTokenIndex = propertyLetStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            var propertySetStmtContext = context as VBAParser.PropertySetStmtContext;
            if (propertySetStmtContext != null)
            {
                lastTokenIndex = propertySetStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            var declareStmtContext = context as VBAParser.DeclareStmtContext;
            if (declareStmtContext != null)
            {
                lastTokenIndex = declareStmtContext.STRINGLITERAL().Last().Symbol.TokenIndex;
                if (declareStmtContext.argList() != null)
                {
                    lastTokenIndex = declareStmtContext.argList().RPAREN().Symbol.TokenIndex;
                }
                if (declareStmtContext.asTypeClause() != null)
                {
                    lastTokenIndex = declareStmtContext.asTypeClause().Stop.TokenIndex;
                }
            }

            var eventStmtContext = context as VBAParser.EventStmtContext;
            if (eventStmtContext != null)
            {
                lastTokenIndex = eventStmtContext.argList().RPAREN().Symbol.TokenIndex;
            }

            return rewriter.GetText(new Interval(firstTokenIndex, lastTokenIndex));
        }
    }
}
