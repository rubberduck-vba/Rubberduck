using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Refactorings.RemoveParameter
{
    class RemoveParameterPresenter
    {
        private readonly Declarations _declarations;
        private readonly Declaration _target;
        private readonly Declaration _method;
        private readonly List<Declaration> _parameters = new List<Declaration>();

        public RemoveParameterPresenter(VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _declarations = parseResult.Declarations;
            
            int indexOfParam;
            string identifierName;

            FindTarget(out _target, out identifierName, selection);
            FindMethod(out _method, out indexOfParam, selection, identifierName);

            if (_target == null && _method != null && indexOfParam != -1)
            {
                var _targets = _declarations.Items
                            .Where(d => d.DeclarationType == DeclarationType.Parameter && d.Scope == _method.Scope)
                            .OrderBy(item => item.Selection.StartLine)
                            .ThenBy(item => item.Selection.StartColumn);

                if (indexOfParam < _targets.Count()) 
                {
                    _target = _targets.ElementAt(indexOfParam); 
                }
                else
                {
                    _target = _targets.ElementAt(_targets.Count() - 1);
                }
            }

            RemoveParameter();
        }

        private void LoadParameters()
        {
            _parameters.AddRange(_declarations.Items
                                    .Where(d => d.DeclarationType == DeclarationType.Parameter 
                                             && d.ComponentName == _method.ComponentName
                                             && d.Project.Equals(_method.Project)
                                             && _method.Context.Start.Line <= d.Selection.StartLine
                                             && _method.Context.Stop.Line >= d.Selection.EndLine
                                             && !(_method.Context.Start.Column > d.Selection.StartColumn && _method.Context.Start.Line == d.Selection.StartLine)
                                             && !(_method.Context.Stop.Column < d.Selection.EndColumn && _method.Context.Stop.Line == d.Selection.EndLine))
                                    .OrderBy(item => item.Selection.StartLine)
                                    .ThenBy(item => item.Selection.StartColumn));
        }

        private void RemoveParameter()
        {
            if (_target == null || _method == null) { return; }

            LoadParameters();

            if (!ConfirmRemove()) { return; }

            AdjustSignatures();
            AdjustReferences(_method.References);
        }

        private bool ConfirmRemove()
        {
            if (IsValidRemove())
            {
                var message = string.Format(RubberduckUI.RemovePresenter_ConfirmParameter, _target.Context.GetText());
                var confirm = MessageBox.Show(message, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (confirm == DialogResult.Yes)
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsValidRemove()
        {
            if (_method.DeclarationType == DeclarationType.PropertyGet &&
                _parameters.FindIndex(item => item.Context.GetText() == _target.Context.GetText()) < 0)
            {
                MessageBox.Show(RubberduckUI.RemoveParamsDialog_RemoveIllegalSetterLetterParameter, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references)
        {
            foreach (var reference in references.Where(item => item.Context != _method.Context))
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
                    var paramToRemove = paramNames.ElementAt(paramIndex);

                    if (!content.Contains(paramToRemove)) { continue; }

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

        private void AdjustSignatures()
        {
            var proc = (dynamic)_method.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = _method.QualifiedName.QualifiedModuleName.Component.CodeModule;
            
            // if we are adjusting a property getter, check if we need to adjust the letter/setter too
            if (_method.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = _declarations.Items.FirstOrDefault(item => item.Scope == _method.Scope &&
                              item.IdentifierName == _method.IdentifierName &&
                              item.DeclarationType == DeclarationType.PropertySet);

                if (setter != null)
                {
                    AdjustSignatures(setter);
                }

                var letter = _declarations.Items.FirstOrDefault(item => item.Scope == _method.Scope &&
                              item.IdentifierName == _method.IdentifierName &&
                              item.DeclarationType == DeclarationType.PropertyLet);

                if (letter != null)
                {
                    AdjustSignatures(letter);
                }
            }
                
            RemoveSignatureParameter(paramList, module);

            foreach (var withEvents in _declarations.Items.Where(item => item.IsWithEvents && item.AsTypeName == _method.ComponentName))
            {
                foreach (var reference in _declarations.FindEventProcedures(withEvents))
                {
                    AdjustSignatures(reference);
                    AdjustReferences(reference.References);
                }
            }

            var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.Project.Equals(_method.Project) &&
                                                               item.IdentifierName == _method.ComponentName + "_" + _method.IdentifierName);
            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustSignatures(interfaceImplentation);
                AdjustReferences(interfaceImplentation.References);
            }
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

            RemoveSignatureParameter(paramList, module);
        }

        private void RemoveSignatureParameter(VBAParser.ArgListContext paramList, CodeModule module)
        {
            for (var lineNum = paramList.Start.Line; lineNum < paramList.Start.Line + paramList.GetSelection().LineCount; lineNum++)
            {
                var content = module.Lines[lineNum, 1];

                if (!content.Contains(_target.Context.GetText())) { continue; }

                var valueToRemove = _target.Context.GetText() != _parameters.Last().Context.GetText() ?
                                    _target.Context.GetText() + "," :
                                    _target.Context.GetText();

                var newContent = content.Replace(valueToRemove, "");

                module.ReplaceLine(lineNum, newContent);
                if (_target.Context.GetText() == _parameters.Last().Context.GetText())
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

                return;
            }
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

            if (method != null && (method.DeclarationType == DeclarationType.PropertySet || method.DeclarationType == DeclarationType.PropertyLet))
            {
                var nonRefMethod = method;

                var getter = _declarations.Items.FirstOrDefault(item => item.ParentScope == nonRefMethod.ParentScope &&
                                              item.IdentifierName == nonRefMethod.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertyGet);

                if (getter != null)
                {
                    method = getter;
                }
            }

            PromptIfTargetImplementsInterface(ref method);
        }

        private void PromptIfTargetImplementsInterface(ref Declaration target)
        {
            var declaration = target;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (target == null || interfaceImplementation == null)
            {
                return;
            }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.RemovePresenter_TargetIsInterfaceMemberImplementation, target.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = MessageBox.Show(message, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirm == DialogResult.No)
            {
                target = null;
                return;
            }

            target = interfaceMember;
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
        }
    }
}
