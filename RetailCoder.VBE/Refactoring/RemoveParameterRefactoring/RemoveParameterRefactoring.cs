using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Refactoring.RemoveParameterRefactoring
{
    public class RemoveParameterRefactoring// : IRefactoring
    {
        private readonly VBProjectParseResult _parseResult;
        private readonly Declarations _declarations;
        private Declaration _target;
        public List<Parameter> Parameters = new List<Parameter>();

        public RemoveParameterRefactoring(VBProjectParseResult parseResult, Declaration target)
        {
            _parseResult = parseResult;
            _declarations = parseResult.Declarations;
            
            if (target.DeclarationType != DeclarationType.Parameter) { throw new ArgumentException("Invalid target type"); }

            FindTarget(out _target, new QualifiedSelection(target.QualifiedName.QualifiedModuleName, target.Selection));

            LoadParameters();
            Parameters.Find(item => Equals(item.Declaration, target)).IsRemoved = true;
        }

        public RemoveParameterRefactoring(VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _parseResult = parseResult;
            _declarations = parseResult.Declarations;

            FindTarget(out _target, selection);

            if (_target != null && (_target.DeclarationType == DeclarationType.PropertySet || _target.DeclarationType == DeclarationType.PropertyLet))
            {
                GetGetter(ref _target);
            }

            LoadParameters();
        }

        public void Refactor()
        {
            PromptIfTargetImplementsInterface(ref _target);
            if (_target == null) { return; }

            RemoveParameters();
        }

        private void PromptIfTargetImplementsInterface(ref Declaration method)
        {
            var declaration = method;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (method == null || interfaceImplementation == null)
            {
                return;
            }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.ReorderPresenter_TargetIsInterfaceMemberImplementation, method.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirm == DialogResult.No)
            {
                method = null;
                return;
            }

            method = interfaceMember;
            var methodParams = FindTargets(method);

            for (var i = 0; i < Parameters.Count; i++)
            {
                Parameters[i] = new Parameter(methodParams.ElementAt(i), Parameters[i].Index, Parameters[i].IsRemoved);
            }
        }

        public Declaration AcquireTarget(QualifiedSelection selection) { return null; }

        private void LoadParameters()
        {
            var index = 0;
            foreach (var target in FindTargets(_target))
            {
                Parameters.Add(new Parameter(target, index++));
            }
        }

        private void RemoveParameters()
        {
            if (_target == null) { throw new NullReferenceException("Parameter is null."); }

            AdjustReferences(_target.References, _target);
            AdjustSignatures();
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references, Declaration method)
        {
            foreach (var reference in references.Where(item => item.Context != method.Context))
            {
                var proc = (dynamic)reference.Context.Parent;
                var module = reference.QualifiedModuleName.Component.CodeModule;

                // This is to prevent throws when this statement fails:
                // (VBAParser.ArgsCallContext)proc.argsCall();
                try
                {
                    var check = (VBAParser.ArgsCallContext)proc.argsCall();
                }
                catch
                {
                    continue;
                }

                var paramList = (VBAParser.ArgsCallContext)proc.argsCall();

                if (paramList == null)
                {
                    continue;
                }

                var numParams = paramList.argCall().Count;  // handles optional variables

                foreach (var param in Parameters.Where(item => item.IsRemoved && item.Index < numParams).Select(item => item.Declaration))
                {
                    RemoveCallParameter(param, paramList, module);
                }
            }
        }

        private void RemoveCallParameter(Declaration paramToRemove, VBAParser.ArgsCallContext paramList, CodeModule module)
        {
            var paramNames = paramList.argCall().Select(arg => arg.GetText()).ToList();
            var paramIndex = Parameters.FindIndex(item => item.Declaration.Context.GetText() == paramToRemove.Context.GetText());

            if (paramIndex >= paramNames.Count) { return; }

            var lineCount = paramList.Stop.Line - paramList.Start.Line + 1; // adjust for total line count

            for (var lineNum = paramList.Start.Line; lineNum < paramList.Start.Line + lineCount; lineNum++)
            {
                var content = module.Lines[lineNum, 1];

                do
                {
                    var paramToRemoveName = paramNames.ElementAt(0).Contains(":=") ? paramNames.Find(item => item.Contains(paramToRemove.IdentifierName + ":=")) : paramNames.ElementAt(paramIndex);

                    if (paramToRemoveName == null || !content.Contains(paramToRemoveName)) { continue; }

                    var valueToRemove = paramToRemoveName != paramNames.Last() ?
                                        paramToRemoveName + "," :
                                        paramToRemoveName;

                    content = content.Replace(valueToRemove, "");

                    module.ReplaceLine(lineNum, content);
                    if (paramToRemoveName == paramNames.Last())
                    {
                        for (var line = lineNum; line >= paramList.Start.Line; line--)
                        {
                            var lineContent = module.Lines[line, 1];
                            if (lineContent.Contains(','))
                            {
                                module.ReplaceLine(line, lineContent.Remove(lineContent.LastIndexOf(','), 1));
                                return;
                            }
                        }
                    }
                } while (paramIndex >= Parameters.Count - 1 && ++paramIndex < paramNames.Count && content.Contains(paramNames.ElementAt(paramIndex)));
            }
        }

        private string GetOldSignature(Declaration target)
        {
            var targetModule = _parseResult.ComponentParseResults.SingleOrDefault(m => m.QualifiedName == target.QualifiedName.QualifiedModuleName);
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

            for (int i = 0, commaCounter = 0; i < signature.Length; i++)
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
            var proc = (dynamic)_target.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = _target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            
            // if we are adjusting a property getter, check if we need to adjust the letter/setter too
            if (_target.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = GetLetterOrSetter(_target, DeclarationType.PropertySet);
                if (setter != null)
                {
                    AdjustSignatures(setter);
                }

                var letter = GetLetterOrSetter(_target, DeclarationType.PropertyLet);
                if (letter != null)
                {
                    AdjustSignatures(letter);
                }
            }
                
            RemoveSignatureParameters(_target, paramList, module);

            foreach (var withEvents in _declarations.Items.Where(item => item.IsWithEvents && item.AsTypeName == _target.ComponentName))
            {
                foreach (var reference in _declarations.FindEventProcedures(withEvents))
                {
                    AdjustReferences(reference.References, reference);
                    AdjustSignatures(reference);
                }
            }

            var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.Project.Equals(_target.Project) &&
                                                               item.IdentifierName == _target.ComponentName + "_" + _target.IdentifierName);
            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustReferences(interfaceImplentation.References, interfaceImplentation);
                AdjustSignatures(interfaceImplentation);
            }
        }

        private Declaration GetLetterOrSetter(Declaration declaration, DeclarationType declarationType)
        {
            return _declarations.Items.FirstOrDefault(item => item.Scope == declaration.Scope &&
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

            var paramsRemoved = Parameters.Where(item => item.IsRemoved).ToList();
            var signature = GetOldSignature(target);

            foreach (var param in paramsRemoved)
            {
                signature = ReplaceCommas(signature.Replace(paramNames.ElementAt(param.Index).GetText(), ""), Parameters.FindIndex(item => item == param) - paramsRemoved.FindIndex(item => item == param));
            }
            var lineNum = paramList.GetSelection().LineCount;

            module.ReplaceLine(paramList.Start.Line, signature);
            module.DeleteLines(paramList.Start.Line + 1, lineNum - 1);
        }

        private IEnumerable<Declaration> FindTargets(Declaration method)
        {
            return _declarations.Items
                              .Where(d => d.DeclarationType == DeclarationType.Parameter
                                       && d.ComponentName == method.ComponentName
                                       && d.Project.Equals(method.Project)
                                       && method.Context.Start.Line <= d.Selection.StartLine
                                       && method.Context.Stop.Line >= d.Selection.EndLine
                                       && !(method.Context.Start.Column > d.Selection.StartColumn && method.Context.Start.Line == d.Selection.StartLine)
                                       && !(method.Context.Stop.Column < d.Selection.EndColumn && method.Context.Stop.Line == d.Selection.EndLine))
                              .OrderBy(item => item.Selection.StartLine)
                              .ThenBy(item => item.Selection.StartColumn);
        }

        /// <summary>
        /// Declaration types that contain parameters that that can be removed.
        /// </summary>
        private static readonly DeclarationType[] ValidDeclarationTypes =
            {
                 DeclarationType.Event,
                 DeclarationType.Function,
                 DeclarationType.Procedure,
                 DeclarationType.PropertyGet,
                 DeclarationType.PropertyLet,
                 DeclarationType.PropertySet
            };

        private void FindTarget(out Declaration target, QualifiedSelection selection)
        {
            target = _declarations.Items
                .Where(item => !item.IsBuiltIn)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item)
                                     || IsSelectedReference(selection, item));

            if (target != null && ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                return;
            }

            target = null;

            var targets = _declarations.Items
                .Where(item => !item.IsBuiltIn
                            && item.ComponentName == selection.QualifiedName.ComponentName
                            && ValidDeclarationTypes.Contains(item.DeclarationType));

            var currentStartLine = 0;
            var currentEndLine = int.MaxValue;
            var currentStartColumn = 0;
            var currentEndColumn = int.MaxValue;

            foreach (var declaration in targets)
            {
                var startLine = declaration.Context.Start.Line;
                var startColumn = declaration.Context.Start.Column;
                var endLine = declaration.Context.Stop.Line;
                var endColumn = declaration.Context.Stop.Column;

                if (startLine <= selection.Selection.StartLine && endLine >= selection.Selection.EndLine &&
                    currentStartLine <= startLine && currentEndLine >= endLine)
                {
                    if (!(startLine == selection.Selection.StartLine && startColumn > selection.Selection.StartColumn ||
                        endLine == selection.Selection.EndLine && endColumn < selection.Selection.EndColumn) &&
                        currentStartColumn <= startColumn && currentEndColumn >= endColumn)
                    {
                        target = declaration;

                        currentStartLine = startLine;
                        currentEndLine = endLine;
                        currentStartColumn = startColumn;
                        currentEndColumn = endColumn;
                    }
                }

                foreach (var reference in declaration.References)
                {
                    var proc = (dynamic)reference.Context.Parent;

                    // This is to prevent throws when this statement fails:
                    // (VBAParser.ArgsCallContext)proc.argsCall();
                    try
                    {
                        var check = (VBAParser.ArgsCallContext)proc.argsCall();
                    }
                    catch
                    {
                        continue;
                    }

                    var paramList = (VBAParser.ArgsCallContext)proc.argsCall();

                    if (paramList == null)
                    {
                        continue;
                    }

                    startLine = paramList.Start.Line;
                    startColumn = paramList.Start.Column;
                    endLine = paramList.Stop.Line;
                    endColumn = paramList.Stop.Column + paramList.Stop.Text.Length + 1;

                    if (startLine <= selection.Selection.StartLine && endLine >= selection.Selection.EndLine &&
                        currentStartLine <= startLine && currentEndLine >= endLine)
                    {
                        if (!(startLine == selection.Selection.StartLine && startColumn > selection.Selection.StartColumn ||
                            endLine == selection.Selection.EndLine && endColumn < selection.Selection.EndColumn) &&
                            currentStartColumn <= startColumn && currentEndColumn >= endColumn)
                        {
                            target = reference.Declaration;

                            currentStartLine = startLine;
                            currentEndLine = endLine;
                            currentStartColumn = startColumn;
                            currentEndColumn = endColumn;
                        }
                    }
                }
            }
        }

        private void GetGetter(ref Declaration method)
        {
            var nonRefMethod = method;

            var getter = _declarations.Items.FirstOrDefault(item => item.Scope == nonRefMethod.Scope &&
                                          item.IdentifierName == nonRefMethod.IdentifierName &&
                                          item.DeclarationType == DeclarationType.PropertyGet);

            if (getter != null)
            {
                method = getter;
            }
        }

        private bool IsSelectedReference(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.References.Any(r =>
                r.QualifiedModuleName == selection.QualifiedName &&
                r.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
        }
    }
}
