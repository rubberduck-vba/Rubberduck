﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    class ReorderParametersPresenter
    {
        private readonly VBE _vbe;
        private readonly IReorderParametersView _view;
        private readonly Declarations _declarations;
        private readonly QualifiedSelection _selection;
        private readonly VBProjectParseResult _parseResult;

        public ReorderParametersPresenter(VBE vbe, IReorderParametersView view, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _vbe = vbe;
            _view = view;
            _view.OkButtonClicked += OnOkButtonClicked;

            _parseResult = parseResult;
            _declarations = parseResult.Declarations;
            _selection = selection;
        }

        public void Show()
        {
            AcquireTarget(_selection);

            if (_view.Target != null)
            {
                LoadParams();

                _view.InitializeParameterGrid();
                _view.ShowDialog();
            }
        }

        private void LoadParams()
        {
            var proc = (dynamic)_view.Target.Context;
            var argList = (VBAParser.ArgListContext)proc.argList();
            var args = argList.arg();

            var index = 0;
            foreach (var arg in args)
            {
                _view.Parameters.Add(new Parameter(arg.ambiguousIdentifier().GetText(), arg.GetText(), index++));
            }
        }

        private void OnOkButtonClicked(object sender, EventArgs e)
        {
            if (!_view.Parameters.Where((t, i) => t.Index != i).Any()) { return; }

            AdjustSignature();
            AdjustReferences();

            foreach (var withEvents in _declarations.Items.Where(item => item.IsWithEvents && item.AsTypeName == _view.Target.ComponentName))
            {
                foreach (var reference in _declarations.FindEventProcedures(withEvents))
                {
                    AdjustSignature(reference);
                }
            }

            var modules = _view.Target.References.GroupBy(r => r.QualifiedModuleName);
        }

        private void AdjustReferences()
        {
            foreach (var reference in _view.Target.References.Where(item => item.Context != _view.Target.Context))
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
                    // update letter/setter methods - needs proper fixing
                    if (reference.Context.Parent.GetText().Contains("Property Let") ||
                        reference.Context.Parent.GetText().Contains("Property Set"))
                    {
                        AdjustSignature(reference);
                    }

                    continue;
                }

                var argList = (VBAParser.ArgsCallContext)proc.argsCall();
                var paramNames = argList.argCall().Select(arg => arg.GetText()).ToList();

                var module = reference.QualifiedModuleName.Component.CodeModule;
                var lineCount = argList.Stop.Line - argList.Start.Line + 1; // adjust for total line count

                var variableIndex = 0;
                for (var line = argList.Start.Line; line < argList.Start.Line + lineCount; line++)
                {
                    var newContent = module.Lines[line, 1];
                    var currentStringIndex = line == argList.Start.Line ? reference.Declaration.IdentifierName.Length : 0;

                    for (var i = variableIndex; i < paramNames.Count; i++)
                    {
                        var variableStringIndex = newContent.IndexOf(paramNames.ElementAt(variableIndex), currentStringIndex);

                        if (variableStringIndex > -1)
                        {
                            var oldVariableString = paramNames.ElementAt(variableIndex);
                            var newVariableString = paramNames.ElementAt(_view.Parameters.IndexOf(_view.Parameters.Find(item => item.Index == variableIndex)));
                            var beginningSub = newContent.Substring(0, variableStringIndex);
                            var replaceSub = newContent.Substring(variableStringIndex).Replace(oldVariableString, newVariableString);

                            newContent = beginningSub + replaceSub;

                            variableIndex++;
                            currentStringIndex = beginningSub.Length + newVariableString.Length;
                        }
                    }

                    module.ReplaceLine(line, newContent);
                }
            }
        }

        // TODO - refactor this
        // Only used for Property Letters/Setters
        // Otherwise, they are caught by the try/catch block used to prevent 
        // value returns from crashing the program
        // Extremely similar to the other AdjustReference
        // One possibility is to create multiple "handler" methods
        // that call another method to do the real work, passing 
        // "module", "argList", and "args" (I believe these are the only
        // ones used
        private void AdjustSignature(IdentifierReference reference)
        {
            var proc = (dynamic)reference.Context.Parent;
            var module = reference.QualifiedModuleName.Component.CodeModule;
            var argList = (VBAParser.ArgListContext)proc.argList();
            var args = argList.arg();

            var variableIndex = 0;
            for (var lineNum = argList.Start.Line; lineNum < argList.Start.Line + argList.GetSelection().LineCount; lineNum++)
            {
                var newContent = module.Lines[lineNum, 1];
                var currentStringIndex = 0;

                for (var i = variableIndex; i < _view.Parameters.Count; i++)
                {
                    var variableStringIndex = newContent.IndexOf(_view.Parameters.Find(item => item.Index == variableIndex).Variable, currentStringIndex);

                    if (variableStringIndex > -1)
                    {
                        var oldVariableString = _view.Parameters.Find(item => item.Index == variableIndex).Variable;
                        var newVariableString = _view.Parameters.ElementAt(i).Variable;
                        var beginningSub = newContent.Substring(0, variableStringIndex);
                        var replaceSub = newContent.Substring(variableStringIndex).Replace(oldVariableString, newVariableString);

                        newContent = beginningSub + replaceSub;

                        variableIndex++;
                        currentStringIndex = beginningSub.Length + newVariableString.Length;
                    }
                }

                module.ReplaceLine(lineNum, newContent);
            }
        }

        private void AdjustSignature(Declaration reference = null)
        {
            var proc = (dynamic)_view.Target.Context;
            var argList = (VBAParser.ArgListContext)proc.argList();
            var module = _view.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;

            if (reference != null)
            {
                proc = (dynamic)reference.Context.Parent;
                module = reference.QualifiedName.QualifiedModuleName.Component.CodeModule;

                if (reference.DeclarationType == DeclarationType.PropertySet)
                {
                    argList = (VBAParser.ArgListContext)proc.children[0].argList();
                }
                else
                {
                    argList = (VBAParser.ArgListContext)proc.subStmt().argList();
                }
            }

            var args = argList.arg();

            if (reference == null && _view.Target.DeclarationType == DeclarationType.PropertyGet)
            {
                var setter = _declarations.Items.Where(item => item.ParentScope == _view.Target.ParentScope &&
                                              item.IdentifierName == _view.Target.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertySet).FirstOrDefault();

                if (setter != null)
                {
                    AdjustSignature(setter);
                }
            }

            var variableIndex = 0;
            for (var lineNum = argList.Start.Line; lineNum < argList.Start.Line + argList.GetSelection().LineCount; lineNum++)
            {
                var newContent = module.Lines[lineNum, 1];
                var currentStringIndex = 0;

                for (var i = variableIndex; i < _view.Parameters.Count; i++)
                {
                    var variableStringIndex = newContent.IndexOf(_view.Parameters.Find(item => item.Index == variableIndex).Variable, currentStringIndex);

                    if (variableStringIndex > -1)
                    {
                        var oldVariableString = _view.Parameters.Find(item => item.Index == variableIndex).Variable;
                        var newVariableString = _view.Parameters.ElementAt(i).Variable;
                        var beginningSub = newContent.Substring(0, variableStringIndex);
                        var replaceSub = newContent.Substring(variableStringIndex).Replace(oldVariableString, newVariableString);

                        newContent = beginningSub + replaceSub;

                        variableIndex++;
                        currentStringIndex = beginningSub.Length + newVariableString.Length;
                    }
                }

                module.ReplaceLine(lineNum, newContent);
            }
        }

        private static readonly DeclarationType[] ValidDeclarationTypes =
            {
                 DeclarationType.Event,
                 DeclarationType.Function,
                 DeclarationType.Procedure,
                 DeclarationType.PropertyGet,
                 DeclarationType.PropertyLet,
                 DeclarationType.PropertySet
            };

        private void AcquireTarget(QualifiedSelection selection)
        {
            var target = _declarations.Items
                .Where(item => !item.IsBuiltIn && ValidDeclarationTypes.Contains(item.DeclarationType))
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item)
                                     || IsSelectedReference(selection, item));

            if (target.DeclarationType == DeclarationType.PropertySet)
            {
                var getter = _declarations.Items.Where(item => item.ParentScope == target.ParentScope &&
                                              item.IdentifierName == target.IdentifierName &&
                                              item.DeclarationType == DeclarationType.PropertyGet).FirstOrDefault();

                if (getter != null)
                {
                    target = getter;
                }
            }

            PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;
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
            var message = string.Format(RubberduckUI.RenamePresenter_TargetIsInterfaceMemberImplementation, target.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = MessageBox.Show(message, RubberduckUI.RenameDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
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
