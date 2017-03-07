using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime.Misc;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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
            {
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
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references, Declaration method)
        {
            foreach (var reference in references.Where(item => item.Context != method.Context))
            {
                var module = reference.QualifiedModuleName.Component.CodeModule;
                {
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
                    RemoveCallParameter(argumentList, module);
                    
                }
            }
        }

        private void RemoveCallParameter(VBAParser.ArgumentListContext paramList, ICodeModule module)
        {
            var paramNames = new List<string>();
            if (paramList.positionalOrNamedArgumentList().positionalArgumentOrMissing() != null)
            {
                paramNames.AddRange(paramList.positionalOrNamedArgumentList().positionalArgumentOrMissing().Select(p =>
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
                paramNames.AddRange(paramList.positionalOrNamedArgumentList().namedArgumentList().namedArgument().Select(p => p.GetText()).ToList());
            }
            if (paramList.positionalOrNamedArgumentList().requiredPositionalArgument() != null)
            {
                paramNames.Add(paramList.positionalOrNamedArgumentList().requiredPositionalArgument().GetText());
            }
            var lineCount = paramList.Stop.Line - paramList.Start.Line + 1; // adjust for total line count

            var newContent = module.GetLines(paramList.Start.Line, lineCount);
            newContent = newContent.Remove(paramList.Start.Column, paramList.GetText().Length);

            var savedParamNames = paramNames;
            for (var index = _model.Parameters.Count - 1; index >= 0; index--)
            {
                var param = _model.Parameters[index];
                if (!param.IsRemoved)
                {
                    continue;
                }

                if (param.Name.Contains("ParamArray"))
                {
                    // handle param arrays
                    while (savedParamNames.Count > index)
                    {
                        savedParamNames.RemoveAt(index);
                    }
                }
                else
                {
                    if (index < savedParamNames.Count && !savedParamNames[index].StripStringLiterals().Contains(":="))
                    {
                        savedParamNames.RemoveAt(index);
                    }
                    else
                    {
                        var paramIndex = savedParamNames.FindIndex(s => s.StartsWith(param.Declaration.IdentifierName + ":="));
                        if (paramIndex != -1 && paramIndex < savedParamNames.Count)
                        {
                            savedParamNames.RemoveAt(paramIndex);
                        }
                    }
                }
            }

            newContent = newContent.Insert(paramList.Start.Column, string.Join(", ", savedParamNames));

            module.ReplaceLine(paramList.Start.Line, newContent.Replace(" _" + Environment.NewLine, string.Empty));
            module.DeleteLines(paramList.Start.Line + 1, lineCount - 1);
        }

        private string GetOldSignature(Declaration target)
        {
            var component = target.QualifiedName.QualifiedModuleName.Component;
            if (component == null)
            {
                throw new InvalidOperationException("Component is null for specified target.");
            }
            var rewriter = _model.State.GetRewriter(component);

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

        private void AdjustSignatures()
        {
            var proc = (dynamic)_model.TargetDeclaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = _model.TargetDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            {
                // if we are adjusting a property getter, check if we need to adjust the letter/setter too
                if (_model.TargetDeclaration.DeclarationType == DeclarationType.PropertyGet)
                {
                    var setter = GetLetterOrSetter(_model.TargetDeclaration, DeclarationType.PropertySet);
                    if (setter != null)
                    {
                        AdjustSignatures(setter);
                        AdjustReferences(setter.References, setter);
                    }

                    var letter = GetLetterOrSetter(_model.TargetDeclaration, DeclarationType.PropertyLet);
                    if (letter != null)
                    {
                        AdjustSignatures(letter);
                        AdjustReferences(letter.References, letter);
                    }
                }

                RemoveSignatureParameters(_model.TargetDeclaration, paramList, module);

                var eventImplementations = _model.Declarations
                    .Where(item => item.IsWithEvents && item.AsTypeName == _model.TargetDeclaration.ComponentName)
                    .SelectMany(withEvents => _model.Declarations.FindEventProcedures(withEvents));

                foreach (var eventImplementation in eventImplementations)
                {
                    AdjustReferences(eventImplementation.References, eventImplementation);
                    AdjustSignatures(eventImplementation);
                }

                var interfaceImplementations = _model.Declarations.FindInterfaceImplementationMembers().Where(item => 
                        item.ProjectId == _model.TargetDeclaration.ProjectId 
                        && item.IdentifierName == _model.TargetDeclaration.ComponentName + "_" + _model.TargetDeclaration.IdentifierName);

                foreach (var interfaceImplentation in interfaceImplementations)
                {
                    AdjustReferences(interfaceImplentation.References, interfaceImplentation);
                    AdjustSignatures(interfaceImplentation);
                }               
            }
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return _model.Declarations.FirstOrDefault(item => item.Scope == declaration.Scope 
                && item.IdentifierName == declaration.IdentifierName 
                && item.DeclarationType == declarationType);
        }

        private void AdjustSignatures(Declaration declaration)
        {
            var proc = (dynamic)declaration.Context.Parent;
            var module = declaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            {
                VBAParser.ArgListContext paramList;

                if (declaration.DeclarationType == DeclarationType.PropertySet
                    || declaration.DeclarationType == DeclarationType.PropertyLet)
                {
                    paramList = (VBAParser.ArgListContext)proc.children[0].argList();
                }
                else
                {
                    paramList = (VBAParser.ArgListContext)proc.subStmt().argList();
                }

                RemoveSignatureParameters(declaration, paramList, module);
            }
        }

        private void RemoveSignatureParameters(Declaration target, VBAParser.ArgListContext paramList, ICodeModule module)
        {
            // property set/let have one more parameter than is listed in the getter parameters
            var nonRemovedParamNames = paramList.arg().Where((a, s) => s >= _model.Parameters.Count || !_model.Parameters[s].IsRemoved).Select(s => s.GetText());
            var signature = GetOldSignature(target);
            signature = signature.Remove(signature.IndexOf('('));
            
            var asTypeText = target.AsTypeContext == null ? string.Empty : " " + target.AsTypeContext.GetText();
            signature += '(' + string.Join(", ", nonRemovedParamNames) + ")" + (asTypeText == " " ? string.Empty : asTypeText);

            var lineCount = paramList.GetSelection().LineCount;
            module.ReplaceLine(paramList.Start.Line, signature.Replace(" _" + Environment.NewLine, string.Empty));
            module.DeleteLines(paramList.Start.Line + 1, lineCount - 1);
        }
    }
}
