using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime;

namespace Rubberduck.UI.Refactorings.RemoveParameter
{
    class RemoveParameterPresenter
    {
        private readonly VBProjectParseResult _parseResult;
        private readonly Declarations _declarations;
        private readonly Declaration _target;
        private readonly Declaration _method;
        private readonly List<Declaration> _parameters = new List<Declaration>();

        public RemoveParameterPresenter(VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _parseResult = parseResult;
            _declarations = parseResult.Declarations;
            
            int indexOfParam;
            string identifierName;

            FindTarget(out _target, out identifierName, selection);
            FindMethod(out _method, out indexOfParam, selection, identifierName);

            if (_method != null && _target == null && indexOfParam != -1)
            {
                var targets = FindTargets(_method).ToList();

                _target = indexOfParam < targets.Count() ? targets.ElementAt(indexOfParam) : targets.ElementAt(targets.Count() - 1);
            }

            if (_method != null && (_method.DeclarationType == DeclarationType.PropertySet || _method.DeclarationType == DeclarationType.PropertyLet))
            {
                GetGetter(out _target, ref _method);
            }

            PromptIfTargetImplementsInterface(ref _target, ref _method);

            RemoveParameter();
        }

        private void LoadParameters()
        {
            _parameters.AddRange(FindTargets(_method));
        }

        private void RemoveParameter()
        {
            if (_target == null || _method == null || !ConfirmRemove()) { return; }

            LoadParameters();

            AdjustReferences(_method.References, _method);
            AdjustSignatures();
        }

        private bool ConfirmRemove()
        {
            var message = string.Format(RubberduckUI.RemovePresenter_ConfirmParameter, _target.Context.GetText());
            var confirm = MessageBox.Show(message, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            
            return confirm == DialogResult.Yes;
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

                RemoveCallParameter(paramList, module);
            }
        }

        private void RemoveCallParameter(VBAParser.ArgsCallContext paramList, CodeModule module)
        {
            var paramNames = paramList.argCall().Select(arg => arg.GetText()).ToList();
            var paramIndex = _parameters.FindIndex(item => item.Context.GetText() == _target.Context.GetText());

            if (paramIndex >= paramNames.Count) { return; }

            var lineCount = paramList.Stop.Line - paramList.Start.Line + 1; // adjust for total line count

            for (var lineNum = paramList.Start.Line; lineNum < paramList.Start.Line + lineCount; lineNum++)
            {
                var content = module.Lines[lineNum, 1];

                do
                {
                    var paramToRemove = paramNames.ElementAt(0).Contains(":=") ? paramNames.Find(item => item.Contains(_target.IdentifierName + ":=")) : paramNames.ElementAt(paramIndex);

                    if (paramToRemove == null || !content.Contains(paramToRemove)) { continue; }

                    var valueToRemove = paramToRemove != paramNames.Last() ?
                                        paramToRemove + "," :
                                        paramToRemove;

                    content = content.Replace(valueToRemove, "");

                    module.ReplaceLine(lineNum, content);
                    if (paramToRemove == paramNames.Last())
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
                } while (paramIndex >= _parameters.Count - 1 && ++paramIndex < paramNames.Count && content.Contains(paramNames.ElementAt(paramIndex)));
            }
        }

        private string GetReplacementSignature(Declaration target)
        {
            var targetModule = _parseResult.ComponentParseResults.SingleOrDefault(m => m.QualifiedName == target.QualifiedName.QualifiedModuleName);
            if (targetModule == null)
            {
                return null;
            }

            var argContext = (VBAParser.ArgContext)target.Context;
            var rewriter = targetModule.GetRewriter();
            rewriter.Replace(argContext.Start.TokenIndex, argContext.Stop.TokenIndex, "");

            // Target.Context is an ArgContext, its parent is an ArgsListContext;
            // the ArgsListContext's parent is the procedure context and it includes the body.
            var context = (ParserRuleContext)target.Context.Parent.Parent;
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

        private string ReplaceCommas(string signature, Declaration target, VBAParser.ArgListContext paramList)
        {
            var parameters = paramList.arg().ToList();

            var indexParamRemoved = parameters.FindIndex(item => item.GetText() == target.Context.GetText());

            if (indexParamRemoved != parameters.Count - 1)
            {
                indexParamRemoved++;
            }

            var commaCounter = 0;
            
            for (int i = 0; i < signature.Length; i++)
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
            var proc = (dynamic)_method.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = _method.QualifiedName.QualifiedModuleName.Component.CodeModule;
            
            // if we are adjusting a property getter, check if we need to adjust the letter/setter too
            if (_method.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = GetLetterOrSetter(_method, DeclarationType.PropertySet);
                if (setter != null)
                {
                    AdjustSignatures(setter);
                }

                var letter = GetLetterOrSetter(_method, DeclarationType.PropertyLet);
                if (letter != null)
                {
                    AdjustSignatures(letter);
                }
            }
                
            RemoveSignatureParameter(_target, paramList, module);

            foreach (var withEvents in _declarations.Items.Where(item => item.IsWithEvents && item.AsTypeName == _method.ComponentName))
            {
                foreach (var reference in _declarations.FindEventProcedures(withEvents))
                {
                    AdjustReferences(reference.References, reference);
                    AdjustSignatures(reference);
                }
            }

            var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.Project.Equals(_method.Project) &&
                                                               item.IdentifierName == _method.ComponentName + "_" + _method.IdentifierName);
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

            var indexOfParam = _parameters.FindIndex(item => item.Context.GetText() == _target.Context.GetText());

            var targets = FindTargets(declaration).ToList();
            var target = indexOfParam < targets.Count() ? targets.ElementAt(indexOfParam) : targets.ElementAt(targets.Count() - 1);

            RemoveSignatureParameter(target, paramList, module);
        }

        private void RemoveSignatureParameter(Declaration target, VBAParser.ArgListContext paramList, CodeModule module)
        {
            var newContent = ReplaceCommas(GetReplacementSignature(target), target, paramList);
            var lineNum = paramList.GetSelection().LineCount;

            module.ReplaceLine(paramList.Start.Line, newContent);
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

        private void FindTarget(out Declaration target, out string identifierName, QualifiedSelection selection)
        {
            target = null;
            identifierName = string.Empty;

            var targets = _declarations.Items
                          .Where(item => item.DeclarationType == DeclarationType.Parameter
                                      && item.ComponentName == selection.QualifiedName.ComponentName
                                      && item.Project.Equals(selection.QualifiedName.Project));

            foreach (var declaration in targets)
            {
                var startLine = declaration.Context.Start.Line;
                var startColumn = declaration.Context.Start.Column;
                var endLine = declaration.Context.Stop.Line;
                var endColumn = declaration.Context.Stop.Column + declaration.Context.Stop.Text.Length + 1;

                if (startLine <= selection.Selection.StartLine && endLine >= selection.Selection.EndLine)
                {
                    if (!(startLine == selection.Selection.StartLine && startColumn > selection.Selection.StartColumn ||
                        endLine == selection.Selection.EndLine && endColumn < selection.Selection.EndColumn))
                    {
                        target = declaration;
                        return;
                    }
                }

                foreach (var reference in declaration.References)
                {
                    startLine = reference.Selection.StartLine;
                    startColumn = reference.Selection.StartColumn;
                    endLine = reference.Selection.EndLine;
                    endColumn = reference.Selection.EndColumn;

                    if (startLine <= selection.Selection.StartLine && endLine >= selection.Selection.EndLine)
                    {
                        if (!(startLine == selection.Selection.StartLine && startColumn > selection.Selection.StartColumn ||
                            endLine == selection.Selection.EndLine && endColumn < selection.Selection.EndColumn))
                        {
                            identifierName = reference.IdentifierName;
                            return;
                        }
                    }
                }
            }
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

        private void FindMethod(out Declaration method, out int indexOfParam, QualifiedSelection selection, string identifierName)
        {
            indexOfParam = -1;

            method = _declarations.Items
                .Where(item => !item.IsBuiltIn)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item));

            if (method != null && ValidDeclarationTypes.Contains(method.DeclarationType))
            {
                return;
            }

            method = null;

            var methods = _declarations.Items
                .Where(item => !item.IsBuiltIn
                            && item.ComponentName == selection.QualifiedName.ComponentName
                            && ValidDeclarationTypes.Contains(item.DeclarationType));

            var currentStartLine = 0;
            var currentEndLine = int.MaxValue;
            var currentStartColumn = 0;
            var currentEndColumn = int.MaxValue;

            foreach (var declaration in methods)
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
                        method = declaration;

                        currentStartLine = startLine;
                        currentEndLine = endLine;
                        currentStartColumn = startColumn;
                        currentEndColumn = endColumn;
                    }
                }

                if (_target == null && identifierName != string.Empty)
                {
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
                                method = reference.Declaration;

                                var args = paramList.argCall().ToList();
                                indexOfParam = args.FindIndex(item => item.GetText() == identifierName);

                                currentStartLine = startLine;
                                currentEndLine = endLine;
                                currentStartColumn = startColumn;
                                currentEndColumn = endColumn;
                            }
                        }
                    }
                }
            }
        }

        private void GetGetter(out Declaration target, ref Declaration method)
        {
            var nonRefMethod = method;

            var getter = _declarations.Items.FirstOrDefault(item => item.Scope == nonRefMethod.Scope &&
                                          item.IdentifierName == nonRefMethod.IdentifierName &&
                                          item.DeclarationType == DeclarationType.PropertyGet);

            if (getter != null)
            {
                method = getter;
            }

            var targets = FindTargets(_method).ToList();
            target = targets.FirstOrDefault(item => _target.IdentifierName == item.IdentifierName);

            if (target == null)
            {
                MessageBox.Show(RubberduckUI.RemoveParamsDialog_RemoveIllegalSetterLetterParameter, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PromptIfTargetImplementsInterface(ref Declaration target, ref Declaration method)
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

            var proc = (dynamic)declaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();

            var indexOfInterfaceParam = paramList.arg().ToList().FindIndex(item => item.GetText() == _target.Context.GetText());
            target = FindTargets(_method).ElementAt(indexOfInterfaceParam);
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
        }
    }
}
