using System;
using System.Collections.Generic;
using System.Linq;
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

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersRefactoring : IRefactoring
    {
        private readonly VBE _vbe;
        private readonly IRefactoringPresenterFactory<IRemoveParametersPresenter> _factory;
        private RemoveParametersModel _model;

        public RemoveParametersRefactoring(VBE vbe, IRefactoringPresenterFactory<IRemoveParametersPresenter> factory)
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

            RemoveParameters();
        }

        public void Refactor(QualifiedSelection target)
        {
            _vbe.ActiveCodePane.CodeModule.SetSelection(target);
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            if (!RemoveParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType) && target.DeclarationType != DeclarationType.Parameter)
            {
                throw new ArgumentException("Invalid declaration type");
            }

            _vbe.ActiveCodePane.CodeModule.SetSelection(target.QualifiedSelection);
            Refactor();
        }

        public void QuickFix(RubberduckParserState parseResult, QualifiedSelection selection)
        {
            _model = new RemoveParametersModel(parseResult, selection, new MessageBox());
            var target = _model.Declarations.FindTarget(selection, new[] { DeclarationType.Parameter });

            // ReSharper disable once PossibleUnintendedReferenceComparison
            _model.Parameters.Find(param => param.Declaration == target).IsRemoved = true;
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
                var proc = (dynamic)reference.Context.Parent;
                var module = reference.QualifiedModuleName.Component.CodeModule;
                VBAParser.ArgsCallContext paramList;

                // This is to prevent throws when this statement fails:
                // (VBAParser.ArgsCallContext)proc.argsCall();
                try { paramList = (VBAParser.ArgsCallContext)proc.argsCall(); }
                catch { continue; }

                if (paramList == null) { continue; }

                RemoveCallParameter(paramList, module);
            }
        }

        private void RemoveCallParameter(VBAParser.ArgsCallContext paramList, CodeModule module)
        {
            var paramNames = paramList.argCall().Select(arg => arg.GetText()).ToList();
            var lineCount = paramList.Stop.Line - paramList.Start.Line + 1; // adjust for total line count

            var newContent = module.Lines[paramList.Start.Line, lineCount].Replace(" _" + Environment.NewLine, string.Empty).RemoveExtraSpacesLeavingIndentation();
            var currentStringIndex = 0;

            foreach (
                var param in
                    _model.Parameters.Where(item => item.IsRemoved && item.Index < paramNames.Count)
                        .Select(item => item.Declaration))
            {
                var paramIndex = _model.Parameters.FindIndex(item => item.Declaration.Context.GetText() == param.Context.GetText()); 
                if (paramIndex >= paramNames.Count) { return; }

                do
                {
                    var paramToRemoveName = paramNames.ElementAt(0).Contains(":=")
                        ? paramNames.Find(item => item.Contains(param.IdentifierName + ":="))
                        : paramNames.ElementAt(paramIndex);

                    if (paramToRemoveName == null || !newContent.Contains(paramToRemoveName))
                    {
                        continue;
                    }

                    var valueToRemove = paramToRemoveName != paramNames.Last()
                        ? paramToRemoveName + ","
                        : paramToRemoveName;

                    var parameterStringIndex = newContent.IndexOf(valueToRemove, currentStringIndex, StringComparison.Ordinal);
                    if (parameterStringIndex <= -1) { continue; }

                    newContent = newContent.Remove(parameterStringIndex, valueToRemove.Length);

                    currentStringIndex = parameterStringIndex;

                    if (paramToRemoveName == paramNames.Last() && newContent.LastIndexOf(',') != -1)
                    {
                        newContent = newContent.Remove(newContent.LastIndexOf(','), 1);
                    }
                } while (paramIndex >= _model.Parameters.Count - 1 && ++paramIndex < paramNames.Count &&
                         newContent.Contains(paramNames.ElementAt(paramIndex)));
            }

            module.ReplaceLine(paramList.Start.Line, newContent);
            module.DeleteLines(paramList.Start.Line + 1, lineCount - 1);
        }

        private string GetOldSignature(Declaration target)
        {
            var module = target.QualifiedName.QualifiedModuleName.Component;
            if (module == null)
            {
                throw new InvalidOperationException("Component is null for specified target.");
            }
            var rewriter = _model.ParseResult.GetRewriter(module);

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

        private string ReplaceCommas(string signature, int indexParamRemoved)
        {
            if (signature.Count(c => c == ',') > indexParamRemoved) { indexParamRemoved++; }

            for (int i = 0, commaCounter = 0; i < signature.Length && indexParamRemoved != 0; i++)
            {
                if (signature.ElementAt(i) == ',')
                {
                    commaCounter++;
                }

                if (commaCounter == indexParamRemoved)
                {
                    return signature.Remove(i, 1);
                }
            }

            return signature;
        }

        private void AdjustSignatures()
        {
            var proc = (dynamic)_model.TargetDeclaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = _model.TargetDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;

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

            var eventImplementations =
                _model.Declarations.Where(
                    item => item.IsWithEvents && item.AsTypeName == _model.TargetDeclaration.ComponentName)
                    .SelectMany(withEvents => _model.Declarations.FindEventProcedures(withEvents));
            foreach (var eventImplementation in eventImplementations)
            {
                AdjustReferences(eventImplementation.References, eventImplementation);
                AdjustSignatures(eventImplementation);
            }

            var interfaceImplementations = _model.Declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.ProjectId == _model.TargetDeclaration.ProjectId &&
                                                               item.IdentifierName == _model.TargetDeclaration.ComponentName + "_" + _model.TargetDeclaration.IdentifierName);
            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustReferences(interfaceImplentation.References, interfaceImplentation);
                AdjustSignatures(interfaceImplentation);
            }
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return _model.Declarations.FirstOrDefault(item => item.Scope == declaration.Scope &&
                              item.IdentifierName == declaration.IdentifierName &&
                              item.DeclarationType == declarationType);
        }

        private void AdjustSignatures(Declaration declaration)
        {
            var proc = (dynamic)declaration.Context.Parent;
            var module = declaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            VBAParser.ArgListContext paramList;

            if (declaration.DeclarationType == DeclarationType.PropertySet ||
                declaration.DeclarationType == DeclarationType.PropertyLet)
            {
                paramList = (VBAParser.ArgListContext)proc.children[0].argList();
            }
            else
            {
                paramList = (VBAParser.ArgListContext)proc.subStmt().argList();
            }

            RemoveSignatureParameters(declaration, paramList, module);
        }

        private void RemoveSignatureParameters(Declaration target, VBAParser.ArgListContext paramList, CodeModule module)
        {
            var paramNames = paramList.arg();

            var paramsRemoved = _model.Parameters.Where(item => item.IsRemoved).ToList();
            var signature = GetOldSignature(target);

            foreach (var param in paramsRemoved)
            {
                try
                {
                    signature = ReplaceCommas(signature.Replace(paramNames.ElementAt(param.Index).GetText(), string.Empty), _model.Parameters.FindIndex(item => item == param) - paramsRemoved.FindIndex(item => item == param));
                }
                catch (ArgumentOutOfRangeException)
                {
                }
            }
            var lineNum = paramList.GetSelection().LineCount;

            module.ReplaceLine(paramList.Start.Line, signature.Replace(" _" + Environment.NewLine, string.Empty));
            module.DeleteLines(paramList.Start.Line + 1, lineNum - 1);
        }
    }
}
