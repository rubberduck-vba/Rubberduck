using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    class ReorderParametersPresenter
    {
        private readonly IReorderParametersView _view;
        private readonly Declarations _declarations;
        private readonly QualifiedSelection _selection;

        public ReorderParametersPresenter(IReorderParametersView view, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _view = view;
            _declarations = parseResult.Declarations;
            _selection = selection;

            _view.OkButtonClicked += OkButtonClicked;
        }

        /// <summary>
        /// Displays the Refactor Parameters dialog window.
        /// </summary>
        public void Show()
        {
            AcquireTarget(_selection);

            if (_view.Target == null) { return; }
            
            LoadParameters();

            if (_view.Parameters.Count < 2) 
            {
                var message = string.Format(RubberduckUI.ReorderPresenter_LessThanTwoParametersError, _view.Target.IdentifierName);
                MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; 
            }

            _view.InitializeParameterGrid();
            _view.ShowDialog();
        }

        /// <summary>
        /// Loads the parameters into the dialog window.
        /// </summary>
        private void LoadParameters()
        {
            var procedure = (dynamic)_view.Target.Context;
            var argList = (VBAParser.ArgListContext)procedure.argList();
            var args = argList.arg();

            var index = 0;
            foreach (var arg in args)
            {
                _view.Parameters.Add(new Parameter(arg.GetText(), index++));
            }
        }

        /// <summary>
        /// Handler for OK button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OkButtonClicked(object sender, EventArgs e)
        {
            if (!_view.Parameters.Where((t, i) => t.Index != i).Any())
            {
                return;
            }

            var indexOfFirstOptionalParam = _view.Parameters.FindIndex(param => param.IsOptional);
            if (indexOfFirstOptionalParam >= 0)
            {
                for (var index = indexOfFirstOptionalParam + 1; index < _view.Parameters.Count; index++)
                {
                    if (!_view.Parameters.ElementAt(index).IsOptional)
                    {
                        MessageBox.Show(RubberduckUI.ReorderPresenter_OptionalParametersMustBeLastError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            var indexOfParamArray = _view.Parameters.FindIndex(param => param.IsParamArray);
            if (indexOfParamArray >= 0)
            {
                if (indexOfParamArray != _view.Parameters.Count - 1)
                {
                    MessageBox.Show(RubberduckUI.ReorderPresenter_ParamArrayError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            AdjustSignatures();
            AdjustReferences(_view.Target.References);
        }

        /// <summary>
        /// Adjusts references to method call.
        /// </summary>
        /// <param name="references">An IEnumberable of IdentifierReference's</param>
        private void AdjustReferences(IEnumerable<IdentifierReference> references)
        {
            foreach (var reference in references.Where(item => item.Context != _view.Target.Context))
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

        /// <summary>
        /// Rewrites method calls.
        /// </summary>
        /// <param name="reference">The reference to the method call to be re-written.</param>
        /// <param name="paramList">The ArgsCallContext of the reference.</param>
        /// <param name="module">The CodeModule to rewrite to.</param>
        private void RewriteCall(IdentifierReference reference, VBAParser.ArgsCallContext paramList, Microsoft.Vbe.Interop.CodeModule module)
        {
            var paramNames = paramList.argCall().Select(arg => arg.GetText()).ToList();

            var lineCount = paramList.Stop.Line - paramList.Start.Line + 1; // adjust for total line count

            var parameterIndex = 0;
            for (var line = paramList.Start.Line; line < paramList.Start.Line + lineCount; line++)
            {
                var newContent = module.Lines[line, 1].Replace(" , ", "");

                var currentStringIndex = line == paramList.Start.Line ? reference.Declaration.IdentifierName.Length : 0;

                for (var i = 0; i < paramNames.Count && parameterIndex < _view.Parameters.Count; i++)
                {
                    var parameterStringIndex = newContent.IndexOf(paramNames.ElementAt(i), currentStringIndex);

                    if (parameterStringIndex > -1)
                    {
                        if (_view.Parameters.ElementAt(parameterIndex).Index >= paramNames.Count)
                        {
                            newContent = newContent.Insert(parameterStringIndex, " , ");
                            i--;
                            parameterIndex++;
                            continue;
                        }

                        var oldParameterString = paramNames.ElementAt(i);
                        var newParameterString = paramNames.ElementAt(_view.Parameters.ElementAt(parameterIndex).Index);
                        var beginningSub = newContent.Substring(0, parameterStringIndex);
                        var replaceSub = newContent.Substring(parameterStringIndex).Replace(oldParameterString, newParameterString);

                        newContent = beginningSub + replaceSub;

                        parameterIndex++;
                        currentStringIndex = beginningSub.Length + newParameterString.Length;
                    }
                }

                module.ReplaceLine(line, newContent);
            }
        }

        /// <summary>
        /// Adjust the signature of a selected method.
        /// Handles setters and letters when a getter is adjusted.
        /// </summary>
        private void AdjustSignatures()
        {
            var proc = (dynamic)_view.Target.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();
            var module = _view.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;

            // if we are reordering a property getter, check if we need to reorder a letter/setter too
            if (_view.Target.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = _declarations.Items.FirstOrDefault(item => item.ParentScope == _view.Target.ParentScope &&
                                              item.IdentifierName == _view.Target.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertySet);

                if (setter != null)
                {
                    AdjustSignatures(setter);
                }

                var letter = _declarations.Items.FirstOrDefault(item => item.ParentScope == _view.Target.ParentScope &&
                              item.IdentifierName == _view.Target.IdentifierName &&
                              item.DeclarationType == DeclarationType.PropertyLet);

                if (letter != null)
                {
                    AdjustSignatures(letter);
                }
            }

            RewriteSignature(paramList, module);

            foreach (var withEvents in _declarations.Items.Where(item => item.IsWithEvents && item.AsTypeName == _view.Target.ComponentName))
            {
                foreach (var reference in _declarations.FindEventProcedures(withEvents))
                {
                    AdjustSignatures(reference);
                    AdjustReferences(reference.References);
                }
            }

            var interfaceImplementations = _declarations.FindInterfaceImplementationMembers()
                                                        .Where(item => item.Project.Equals(_view.Target.Project) &&
                                                               item.IdentifierName == _view.Target.ComponentName + "_" + _view.Target.IdentifierName);
            foreach (var interfaceImplentation in interfaceImplementations)
            {
                AdjustSignatures(interfaceImplentation);
                AdjustReferences(interfaceImplentation.References);
            }
        }

        /// <summary>
        /// Adjust the signature of a reference to a given method.
        /// Used for letters.
        /// </summary>
        /// <param name="reference">A reference to the method signature to adjust.</param>
        private void AdjustSignatures(IdentifierReference reference)
        {
            var proc = (dynamic)reference.Context.Parent;
            var module = reference.QualifiedModuleName.Component.CodeModule;
            var paramList = (VBAParser.ArgListContext)proc.argList();

            RewriteSignature(paramList, module);
        }

        /// <summary>
        /// Adjust the signature of a declaration of a given method.
        /// </summary>
        /// <param name="declaration">A Declaration of the method signature to adjust.</param>
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

            RewriteSignature(paramList, module);
        }

        /// <summary>
        /// Rewrites the signature of a given method.
        /// </summary>
        /// <param name="paramList">The ArgListContext of the method signature being adjusted.</param>
        /// <param name="module">The CodeModule of the method signature being adjusted.</param>
        private void RewriteSignature(VBAParser.ArgListContext paramList, Microsoft.Vbe.Interop.CodeModule module)
        {
            var args = paramList.arg();

            var parameterIndex = 0;
            for (var lineNum = paramList.Start.Line; lineNum < paramList.Start.Line + paramList.GetSelection().LineCount; lineNum++)
            {
                var newContent = module.Lines[lineNum, 1];
                var currentStringIndex = 0;

                for (var i = parameterIndex; i < _view.Parameters.Count; i++)
                {
                    var parameterStringIndex = newContent.IndexOf(_view.Parameters.Find(item => item.Index == parameterIndex).FullDeclaration, currentStringIndex);

                    if (parameterStringIndex > -1)
                    {
                        var oldVariableString = _view.Parameters.Find(item => item.Index == parameterIndex).FullDeclaration;
                        var newVariableString = _view.Parameters.ElementAt(i).FullDeclaration;
                        var beginningSub = newContent.Substring(0, parameterStringIndex);
                        var replaceSub = newContent.Substring(parameterStringIndex).Replace(oldVariableString, newVariableString);

                        newContent = beginningSub + replaceSub;

                        parameterIndex++;
                        currentStringIndex = beginningSub.Length + newVariableString.Length;
                    }
                }

                module.ReplaceLine(lineNum, newContent);
            }
        }

        /// <summary>
        /// Declaration types that contain parameters that that can be adjusted.
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

        /// <summary>
        /// Get the target Declaration to adjust.
        /// </summary>
        /// <param name="selection">The user selection specifying which method signature to adjust.</param>
        private void AcquireTarget(QualifiedSelection selection)
        {
            Declaration target;

            FindTarget(out target, selection);

            if (target != null && target.DeclarationType == DeclarationType.PropertySet)
            {
                var getter = _declarations.Items.FirstOrDefault(item => item.ParentScope == target.ParentScope &&
                                              item.IdentifierName == target.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertyGet);

                if (getter != null)
                {
                    target = getter;
                }
            }

            PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;
        }

        /// <summary>
        /// Gets the target to adjust given a selection.
        /// </summary>
        /// <param name="target">The value to place the target in.</param>
        /// <param name="selection">The user selection specifying what method signature to adjust.</param>
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

        /// <summary>
        /// Displays a prompt asking the user whether the method signature should be adjusted
        /// if the target declaration implements an interface method.
        /// </summary>
        /// <param name="target">The target declaration.</param>
        private void PromptIfTargetImplementsInterface(ref Declaration target)
        {
            var declaration = target;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (target == null || interfaceImplementation == null)
            {
                return;
            }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.ReorderPresenter_TargetIsInterfaceMemberImplementation, target.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
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
