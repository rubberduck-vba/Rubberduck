using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Refactorings.RemoveParameter
{
    class RemoveParameterPresenter
    {
        private readonly Declarations _declarations;
        private readonly QualifiedSelection _selection;
        private readonly Declaration _target;
        private readonly Declaration _method;
        private readonly List<Parameter> _parameters = new List<Parameter>();

        public RemoveParameterPresenter(VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _declarations = parseResult.Declarations;
            _selection = selection;

            FindTarget(out _target, selection);

            if (_target == null) { return; }
            FindMethod(out _method, selection);

            RemoveParameter();
        }

        public RemoveParameterPresenter(Declaration target)
        {
            if (target == null)
            {
                return;
            }

            if (target.DeclarationType != DeclarationType.Parameter)
            {
                throw new ArgumentException("Expected DeclarationType.Parameter, received DeclarationType." + target.DeclarationType.ToString() + ".");
            }

            _target = target;
            if (_target == null) { return; }

            //FindMethod(out _method, selection);

            RemoveParameter();
        }

        private void LoadParameters()
        {
            var procedure = (dynamic)_target.Context.Parent;
            var argList = (VBAParser.ArgListContext)procedure;
            var args = argList.arg();

            var index = 0;
            foreach (var arg in args)
            {
                _parameters.Add(new Parameter(arg.GetText(), index++));
            }
        }

        private void RemoveParameter()
        {
            LoadParameters();

            AdjustSignatures();
            AdjustReferences(_method.References);
        }

        private void AdjustReferences(IEnumerable<IdentifierReference> references)
        {
            foreach (var reference in references.Where(item => item.Context != _target.Context))
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

                RewriteCall(reference, paramList, module);
            }
        }

        private void RewriteCall(IdentifierReference reference, VBAParser.ArgsCallContext paramList, Microsoft.Vbe.Interop.CodeModule module)
        {
            var paramNames = paramList.argCall().Select(arg => arg.GetText()).ToList();
            var paramIndex = _parameters.FindIndex(item => item.FullDeclaration == _target.Context.GetText());

            if (paramIndex >= paramNames.Count) { return; }

            var paramToRemove = paramNames.ElementAt(paramIndex);

            var lineCount = paramList.Stop.Line - paramList.Start.Line + 1; // adjust for total line count

            for (var lineNum = paramList.Start.Line; lineNum < paramList.Start.Line + lineCount; lineNum++)
            {
                var content = module.Lines[lineNum, 1];

                if (!content.Contains(paramToRemove)) { continue; }

                var valueToRemove = paramToRemove != paramNames.Last() ?
                                    paramToRemove + "," :
                                    paramToRemove;

                var newContent = content.Replace(valueToRemove, "");

                module.ReplaceLine(lineNum, newContent);
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

                return;
            }
        }

        private void AdjustSignatures()
        {
            var proc = (dynamic)_target.Context.Parent;
            var paramList = (VBAParser.ArgListContext)proc;
            var module = _target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            
            // if we are reordering a property getter, check if we need to reorder a letter/setter too
            if (_target.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = _declarations.Items.FirstOrDefault(item => item.ParentScope == _target.ParentScope &&
                                              item.IdentifierName == _target.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertySet);

                if (setter != null)
                {
                    AdjustSignatures(setter);
                }

                var letter = _declarations.Items.FirstOrDefault(item => item.ParentScope == _target.ParentScope &&
                              item.IdentifierName == _target.IdentifierName &&
                              item.DeclarationType == DeclarationType.PropertyLet);

                if (letter != null)
                {
                    AdjustSignatures(letter);
                }
            }

            RemoveSignatureParameter(paramList, module);

            foreach (var withEvents in _declarations.Items.Where(item => item.IsWithEvents && item.AsTypeName == _target.ComponentName))
            {
                foreach (var reference in _declarations.FindEventProcedures(withEvents))
                {
                    AdjustSignatures(reference);
                }
            }

            var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.Project.Equals(_target.Project) &&
                                                               item.IdentifierName == _target.ComponentName + "_" + _target.IdentifierName);
            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustSignatures(interfaceImplentation);

                AdjustReferences(interfaceImplentation.References);
            }
        }

        private void AdjustSignatures(IdentifierReference reference)
        {
            var proc = (dynamic)reference.Context.Parent;
            var module = reference.QualifiedModuleName.Component.CodeModule;
            var paramList = (VBAParser.ArgListContext)proc.argList();

            RemoveSignatureParameter(paramList, module);
        }

        private void AdjustSignatures(Declaration declaration)
        {
            var proc = (dynamic)_target.Context.Parent;
            var paramList = (VBAParser.ArgListContext)proc;
            var module = _target.QualifiedName.QualifiedModuleName.Component.CodeModule;

            RemoveSignatureParameter(paramList, module);
        }

        private void RemoveSignatureParameter(VBAParser.ArgListContext paramList, Microsoft.Vbe.Interop.CodeModule module)
        {
            for (var lineNum = paramList.Start.Line; lineNum < paramList.Start.Line + paramList.GetSelection().LineCount; lineNum++)
            {
                var content = module.Lines[lineNum, 1];

                if (!content.Contains(_target.Context.GetText())) { continue; }

                var valueToRemove = _target.Context.GetText() != _parameters.Last().FullDeclaration ?
                                    _target.Context.GetText() + "," :
                                    _target.Context.GetText();

                var newContent = content.Replace(valueToRemove, "");

                module.ReplaceLine(lineNum, newContent);
                if (_target.Context.GetText() == _parameters.Last().FullDeclaration)
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

        private void FindTarget(out Declaration target, QualifiedSelection selection)
        {
            target = null;

            var targets = _declarations.Items
                        .Where(item => item.DeclarationType == DeclarationType.Parameter
                                    && item.ComponentName == selection.QualifiedName.ComponentName
                                    && item.Project.Equals(selection.QualifiedName.Project));

            if (targets == null)
            {
                return;
            }

            var currentStartLine = 0;
            var currentEndLine = int.MaxValue;
            var currentStartColumn = 0;
            var currentEndColumn = int.MaxValue;

            foreach (var declaration in targets)
            {
                var startLine = declaration.Context.Start.Line;
                var startColumn = declaration.Context.Start.Column;
                var endLine = declaration.Context.Stop.Line;
                var endColumn = declaration.Context.Stop.Column + declaration.Context.Stop.Text.Length + 1;

                var d = declaration.Context.GetSelection();

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

            if (target == null) { return; }

            var message = string.Format(RubberduckUI.RemovePresenter_ConfirmParameter, target.Context.GetText());
            var confirm = MessageBox.Show(message, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirm == DialogResult.No)
            {
                target = null;
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

        private void FindMethod(out Declaration method, QualifiedSelection selection)
        {
            method = _declarations.Items
                .Where(item => !item.IsBuiltIn)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item));

            if (method != null && ValidDeclarationTypes.Contains(method.DeclarationType))
            {
                return;
            }

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
                var startLine = declaration.Context.GetSelection().StartLine;
                var startColumn = declaration.Context.GetSelection().StartColumn;
                var endLine = declaration.Context.GetSelection().EndLine;
                var endColumn = declaration.Context.GetSelection().EndColumn;

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
