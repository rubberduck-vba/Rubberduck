using Antlr4.Runtime.Misc;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersRefactoring : IRefactoring
    {
        private readonly VBE _vbe;
        private readonly IRefactoringPresenterFactory<IReorderParametersPresenter> _factory;
        private ReorderParametersModel _model;
        private readonly IMessageBox _messageBox;

        public ReorderParametersRefactoring(VBE vbe, IRefactoringPresenterFactory<IReorderParametersPresenter> factory, IMessageBox messageBox)
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

            AdjustReferences(_model.TargetDeclaration.References);
            AdjustSignatures();

            _model.State.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            _vbe.ActiveCodePane.CodeModule.SetSelection(target);
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            if (!ReorderParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                throw new ArgumentException("Invalid declaration type");
            }

            _vbe.ActiveCodePane.CodeModule.SetSelection(target.QualifiedSelection);
            Refactor();
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
                dynamic proc = reference.Context;
                var module = reference.QualifiedModuleName.Component.CodeModule;
                VBAParser.ArgumentListContext paramList = null;

                var callStmt = ParserRuleContextHelper.GetParent<VBAParser.CallStmtContext>(reference.Context);
                if (callStmt != null)
                {
                    paramList = callStmt.argumentList();
                }
                else if (reference.Context is VBAParser.IndexExprContext)
                {
                    paramList = ((VBAParser.IndexExprContext)reference.Context).argumentList();
                }

                if (paramList == null) { continue; }

                RewriteCall(paramList, module);
            }
        }

        private void RewriteCall(VBAParser.ArgumentListContext paramList, CodeModule module)
        {
            List<string> paramNames = new List<string>();
            if (paramList.positionalOrNamedArgumentList().positionalArgumentOrMissing() != null)
            {
                paramNames.AddRange(paramList.positionalOrNamedArgumentList().positionalArgumentOrMissing().Select(p =>
                {
                    if (p is VBAParser.SpecifiedPositionalArgumentContext)
                    {
                        return ((VBAParser.SpecifiedPositionalArgumentContext)p).positionalArgument().GetText();
                    }
                    else
                    {
                        return string.Empty;
                    }
                }).ToList());
            }
            if (paramList.positionalOrNamedArgumentList().namedArgumentList() != null)
            {
                paramNames.AddRange(paramList.positionalOrNamedArgumentList().namedArgumentList().namedArgument().Select(p => p.GetText()).ToList());
            }
            if (paramList.positionalOrNamedArgumentList().requiredPositionalArgument() != null)
            {
                paramNames.Add(paramList.positionalOrNamedArgumentList().requiredPositionalArgument().GetText());
            }

            var lineCount = paramList.Stop.Line - paramList.Start.Line + 1; // adjust for total line count

            var newContent = module.Lines[paramList.Start.Line, lineCount].Replace(" _" + Environment.NewLine, string.Empty).RemoveExtraSpacesLeavingIndentation();

            var parameterIndex = 0;
            var currentStringIndex = 0;

            for (var i = 0; i < paramNames.Count && parameterIndex < _model.Parameters.Count; i++)
            {
                var parameterStringIndex = newContent.IndexOf(paramNames.ElementAt(i), currentStringIndex, StringComparison.Ordinal);

                if (parameterStringIndex <= -1) { continue; }

                var oldParameterString = paramNames.ElementAt(i);
                var newParameterString = paramNames.ElementAt(_model.Parameters.ElementAt(parameterIndex).Index);
                var beginningSub = newContent.Substring(0, parameterStringIndex);
                var replaceSub = newContent.Substring(parameterStringIndex).Replace(oldParameterString, newParameterString);

                newContent = beginningSub + replaceSub;

                parameterIndex++;
                currentStringIndex = beginningSub.Length + newParameterString.Length;
            }

            module.ReplaceLine(paramList.Start.Line, newContent);
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

        private void RewriteSignature(Declaration target, VBAParser.ArgListContext paramList, CodeModule module)
        {
            var argList = paramList.arg();

            var newContent = GetOldSignature(target);
            var lineNum = paramList.GetSelection().LineCount;

            var parameterIndex = 0;
            var currentStringIndex = 0;

            for (var i = parameterIndex; i < _model.Parameters.Count; i++)
            {
                var oldParam = argList.ElementAt(parameterIndex).GetText();
                var newParam = argList.ElementAt(_model.Parameters.ElementAt(parameterIndex).Index).GetText();
                var parameterStringIndex = newContent.IndexOf(oldParam, currentStringIndex, StringComparison.Ordinal);

                if (parameterStringIndex > -1)
                {
                    var beginningSub = newContent.Substring(0, parameterStringIndex);
                    var replaceSub = newContent.Substring(parameterStringIndex).Replace(oldParam, newParam);

                    newContent = beginningSub + replaceSub;

                    parameterIndex++;
                    currentStringIndex = beginningSub.Length + newParam.Length;
                }
            }

            module.ReplaceLine(paramList.Start.Line, newContent.Replace(" _" + Environment.NewLine, string.Empty));
            module.DeleteLines(paramList.Start.Line + 1, lineNum - 1);
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
