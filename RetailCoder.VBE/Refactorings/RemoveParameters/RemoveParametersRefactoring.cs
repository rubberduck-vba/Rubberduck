using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersRefactoring : IRefactoring
    {
        private readonly IRefactoringPresenterFactory<RemoveParametersPresenter> _factory;
        private RemoveParametersModel _model;

        public RemoveParametersRefactoring(IRefactoringPresenterFactory<RemoveParametersPresenter> factory)
        {
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
            target.Select();
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            if (!RemoveParametersModel.ValidDeclarationTypes.Contains(target.DeclarationType) && target.DeclarationType != DeclarationType.Parameter)
            {
                throw new ArgumentException("Invalid declaration type");
            }

            target.QualifiedSelection.Select();
            Refactor();
        }

        public void QuickFix(VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _model = new RemoveParametersModel(parseResult, selection);
            var target = _model.FindTarget(selection, new[] { DeclarationType.Parameter });

            // ReSharper disable once PossibleUnintendedReferenceComparison
            _model.Parameters.Find(param => param.Declaration == target).IsRemoved = true;
            RemoveParameters();
        }

        private void RemoveParameters()
        {
            if (_model.TargetDeclaration == null) { throw new NullReferenceException("Parameter is null."); }

            AdjustReferences(_model.TargetDeclaration.References.OrderByDescending(item => item.Selection.StartLine), _model.TargetDeclaration);
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
                var numParams = paramList.argCall().Count;  // handles optional variables

                foreach (var param in _model.Parameters.Where(item => item.IsRemoved && item.Index < numParams).Select(item => item.Declaration))
                {
                    RemoveCallParameter(param, paramList, module);
                }
            }
        }

        private void RemoveCallParameter(Declaration paramToRemove, VBAParser.ArgsCallContext paramList, CodeModule module)
        {
            var paramNames = paramList.argCall().Select(arg => arg.GetText()).ToList();
            var paramIndex = _model.Parameters.FindIndex(item => item.Declaration.Context.GetText() == paramToRemove.Context.GetText());

            if (paramIndex >= paramNames.Count) { return; }

            var lineCount = paramList.Stop.Line - paramList.Start.Line + 1; // adjust for total line count

            var newContent = module.Lines[paramList.Start.Line, lineCount].Replace(" _", "").RemoveExtraSpaces();

            do
            {
                var paramToRemoveName = paramNames.ElementAt(0).Contains(":=") ? paramNames.Find(item => item.Contains(paramToRemove.IdentifierName + ":=")) : paramNames.ElementAt(paramIndex);

                if (paramToRemoveName == null || !newContent.Contains(paramToRemoveName)) { continue; }

                var valueToRemove = paramToRemoveName != paramNames.Last() ?
                                    paramToRemoveName + "," :
                                    paramToRemoveName;

                newContent = newContent.Replace(valueToRemove, "");

                if (paramToRemoveName == paramNames.Last() && newContent.LastIndexOf(',') != -1)
                {
                    newContent = newContent.Remove(newContent.LastIndexOf(','), 1);
                }
            } while (paramIndex >= _model.Parameters.Count - 1 && ++paramIndex < paramNames.Count && newContent.Contains(paramNames.ElementAt(paramIndex)));

            module.ReplaceLine(paramList.Start.Line, newContent);
            module.DeleteLines(paramList.Start.Line + 1, lineCount - 1);
        }

        private string GetOldSignature(Declaration target)
        {
            var targetModule = _model.ParseResult.ComponentParseResults.SingleOrDefault(m => m.QualifiedName == target.QualifiedName.QualifiedModuleName);
            if (targetModule == null)
            {
                return null;
            }

            var rewriter = targetModule.GetRewriter();

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
                }

                var letter = GetLetterOrSetter(_model.TargetDeclaration, DeclarationType.PropertyLet);
                if (letter != null)
                {
                    AdjustSignatures(letter);
                }
            }

            RemoveSignatureParameters(_model.TargetDeclaration, paramList, module);

            foreach (var withEvents in _model.Declarations.Items.Where(item => item.IsWithEvents && item.AsTypeName == _model.TargetDeclaration.ComponentName))
            {
                foreach (var reference in _model.Declarations.FindEventProcedures(withEvents))
                {
                    AdjustReferences(reference.References.OrderByDescending(item => item.Selection.StartLine), reference);
                    AdjustSignatures(reference);
                }
            }

            var interfaceImplementations = _model.Declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.Project.Equals(_model.TargetDeclaration.Project) &&
                                                               item.IdentifierName == _model.TargetDeclaration.ComponentName + "_" + _model.TargetDeclaration.IdentifierName);
            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustReferences(interfaceImplentation.References.OrderByDescending(item => item.Selection.StartLine), interfaceImplentation);
                AdjustSignatures(interfaceImplentation);
            }
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return _model.Declarations.Items.FirstOrDefault(item => item.Scope == declaration.Scope &&
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
                    signature = ReplaceCommas(signature.Replace(paramNames.ElementAt(param.Index).GetText(), ""), _model.Parameters.FindIndex(item => item == param) - paramsRemoved.FindIndex(item => item == param));
                }
                catch (ArgumentOutOfRangeException)
                {
                }
            }
            var lineNum = paramList.GetSelection().LineCount;

            module.ReplaceLine(paramList.Start.Line, signature);
            module.DeleteLines(paramList.Start.Line + 1, lineNum - 1);
        }
    }
}
