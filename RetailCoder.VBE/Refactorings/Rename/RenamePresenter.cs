﻿using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenter
    {
        private readonly IRenameView _view;
        private readonly RenameModel _model;

        public RenamePresenter(IRenameView view, RenameModel model)
        {
            _view = view;
            _view.OkButtonClicked += OnViewOkButtonClicked;

            _model = model;
        }

        public RenameModel Show()
        {
            if (_model.Target != null)
            {
                _view.Target = _model.Target;
                _view.ShowDialog();
            }

            return _model;
        }

        public RenameModel Show(Declaration target)
        {
            _model.PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;
            _view.ShowDialog();
            return _model;
        }

/*        private Declaration AmbiguousId()
        {
            var values = _model.Declarations.Items.Where(item => (item.Scope.Contains(_model.Target.Scope)
                                              || _model.Target.ParentScope.Contains(item.ParentScope))
                                              && _view.NewName == item.IdentifierName);

            if (values.Any())
            {
                return values.FirstOrDefault();
            }

            foreach (var reference in _model.Target.References)
            {
                var potentialDeclarations = _model.Declarations.Items.Where(item => !item.IsBuiltIn
                                                         && item.Project.Equals(reference.Declaration.Project)
                                                         && ((item.Context != null
                                                         && item.Context.Start.Line <= reference.Selection.StartLine
                                                         && item.Context.Stop.Line >= reference.Selection.EndLine)
                                                         || (item.Selection.StartLine <= reference.Selection.StartLine
                                                         && item.Selection.EndLine >= reference.Selection.EndLine))
                                                         && item.QualifiedName.QualifiedModuleName.ComponentName == reference.QualifiedModuleName.ComponentName);

                var currentSelection = new Selection(0, 0, int.MaxValue, int.MaxValue);

                Declaration target = null;
                foreach (var item in potentialDeclarations)
                {
                    var startLine = item.Context == null ? item.Selection.StartLine : item.Context.Start.Line;
                    var endLine = item.Context == null ? item.Selection.EndLine : item.Context.Stop.Column;
                    var startColumn = item.Context == null ? item.Selection.StartColumn : item.Context.Start.Column;
                    var endColumn = item.Context == null ? item.Selection.EndColumn : item.Context.Stop.Column;

                    var selection = new Selection(startLine, startColumn, endLine, endColumn);

                    if (currentSelection.Contains(selection))
                    {
                        currentSelection = selection;
                        target = item;
                    }
                }

                if (target == null) { continue; }

                values = _model.Declarations.Items.Where(item => (item.Scope.Contains(target.Scope)
                                              || target.ParentScope.Contains(item.ParentScope))
                                              && _view.NewName == item.IdentifierName);

                if (values.Any())
                {
                    return values.FirstOrDefault();
                }
            }

            return null;
        }

        private static readonly DeclarationType[] ModuleDeclarationTypes =
        {
            DeclarationType.Class,
            DeclarationType.Module
        };*/

        public event EventHandler<string> OkButtonClicked;
        protected virtual void OnOkButtonClicked(string e)
        {
            var handler = OkButtonClicked;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private void OnViewOkButtonClicked(object sender, EventArgs e)
        {
            OnOkButtonClicked(_view.NewName);
        }

        /*private void RenameModule()
        {
            try
            {
                var module = _model.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                if (module != null)
                {
                    if (module.Parent.Type == vbext_ComponentType.vbext_ct_Document)
                    {
                        module.Parent.Properties.Item("_CodeName").Value = _view.NewName;
                    }
                    else if (module.Parent.Type == vbext_ComponentType.vbext_ct_MSForm)
                    {
                        var codeModule = (CodeModuleClass)module;
                        codeModule.Parent.Name = _view.NewName;
                        module.Parent.Properties.Item("Caption").Value = _view.NewName;
                    }
                    else
                    {
                        module.Name = _view.NewName;
                    }
                }
            }
            catch (COMException)
            {
                MessageBox.Show(RubberduckUI.RenameDialog_ModuleRenameError, RubberduckUI.RenameDialog_Caption);
            }
        }

        private void RenameProject()
        {
            try
            {
                var project = _model.VBE.VBProjects.Cast<VBProject>().FirstOrDefault(p => p.Name == _model.Target.IdentifierName);
                if (project != null)
                {
                    project.Name = _view.NewName;
                }
            }
            catch (COMException)
            {
                MessageBox.Show(RubberduckUI.RenameDialog_ProjectRenameError, RubberduckUI.RenameDialog_Caption);
            }
        }

        private void RenameDeclaration()
        {
            if (_model.Target.DeclarationType == DeclarationType.Control)
            {
                RenameControl();
                return;
            }

            var module = _model.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            var newContent = GetReplacementLine(module, _model.Target, _view.NewName);

            if (_model.Target.DeclarationType == DeclarationType.Parameter)
            {
                var argList = (VBAParser.ArgListContext)_model.Target.Context.Parent;
                var lineNum = argList.GetSelection().LineCount;

                module.ReplaceLine(argList.Start.Line, newContent);
                module.DeleteLines(argList.Start.Line + 1, lineNum - 1);
            }
            else if (!_model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                module.ReplaceLine(_model.Target.Selection.StartLine, newContent);
            }
            else
            {
                var members = _model.Declarations[_model.Target.IdentifierName]
                    .Where(item => item.Project == _model.Target.Project
                        && item.ComponentName == _model.Target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property));

                foreach (var member in members)
                {
                    newContent = GetReplacementLine(module, member, _view.NewName);
                    module.ReplaceLine(member.Selection.StartLine, newContent);
                }
            }
        }

        private void RenameControl()
        {
            try
            {
                var form = _model.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                var control = ((dynamic)form.Parent.Designer).Controls(_model.Target.IdentifierName);

                foreach (var handler in _model.Declarations.FindEventHandlers(_model.Target).OrderByDescending(h => h.Selection.StartColumn))
                {
                    var newMemberName = handler.IdentifierName.Replace(control.Name + '_', _view.NewName + '_');
                    var module = handler.Project.VBComponents.Item(handler.ComponentName).CodeModule;

                    var content = module.Lines[handler.Selection.StartLine, 1];
                    var newContent = GetReplacementLine(content, newMemberName, handler.Selection);
                    module.ReplaceLine(handler.Selection.StartLine, newContent);
                }

                control.Name = _view.NewName;
            }
            catch (RuntimeBinderException)
            {
            }
            catch (COMException)
            {
            }
        }

        private void RenameUsages(Declaration target, string interfaceName = null)
        {
            // todo: refactor

            // rename interface member
            if (_model.Declarations.FindInterfaceMembers().Contains(target))
            {
                var implementations = _model.Declarations.FindInterfaceImplementationMembers()
                    .Where(m => m.IdentifierName == target.ComponentName + '_' + target.IdentifierName);

                foreach (var member in implementations.OrderByDescending(m => m.Selection.StartColumn))
                {
                    try
                    {
                        var newMemberName = target.ComponentName + '_' + _view.NewName;
                        var module = member.Project.VBComponents.Item(member.ComponentName).CodeModule;

                        var content = module.Lines[member.Selection.StartLine, 1];
                        var newContent = GetReplacementLine(content, newMemberName, member.Selection);
                        module.ReplaceLine(member.Selection.StartLine, newContent);
                        RenameUsages(member, target.ComponentName);
                    }
                    catch (COMException)
                    {
                        // gulp
                    }
                }

                return;
            }

            var modules = target.References.GroupBy(r => r.QualifiedModuleName);
            foreach (var grouping in modules)
            {
                var module = grouping.Key.Component.CodeModule;
                foreach (var line in grouping.GroupBy(reference => reference.Selection.StartLine))
                {
                    foreach (var reference in line.OrderByDescending(l => l.Selection.StartColumn))
                    {
                        var content = module.Lines[line.Key, 1];
                        string newContent;

                        if (interfaceName == null)
                        {
                            newContent = GetReplacementLine(content, _view.NewName,
                                reference.Selection);
                        }
                        else
                        {
                            newContent = GetReplacementLine(content,
                                interfaceName + "_" + _view.NewName,
                                reference.Selection);
                        }

                        module.ReplaceLine(line.Key, newContent);
                    }
                }

                // renaming interface
                if (grouping.Any(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext))
                {
                    var members = _model.Declarations.FindMembers(target).OrderByDescending(m => m.Selection.StartColumn);
                    foreach (var member in members)
                    {
                        var oldMemberName = target.IdentifierName + '_' + member.IdentifierName;
                        var newMemberName = _view.NewName + '_' + member.IdentifierName;
                        var method = _model.Declarations[oldMemberName].SingleOrDefault(m => m.QualifiedName.QualifiedModuleName == grouping.Key);
                        if (method == null)
                        {
                            continue;
                        }

                        var content = module.Lines[method.Selection.StartLine, 1];
                        var newContent = GetReplacementLine(content, newMemberName, member.Selection);
                        module.ReplaceLine(method.Selection.StartLine, newContent);
                    }
                }
            }
        }

        private string GetReplacementLine(string content, string newName, Selection selection)
        {
            var contentWithoutOldName = content.Remove(selection.StartColumn - 1, selection.EndColumn - selection.StartColumn);
            return contentWithoutOldName.Insert(selection.StartColumn - 1, newName);
        }

        private string GetReplacementLine(CodeModule module, Declaration target, string newName)
        {
            var targetModule = _model.ParseResult.ComponentParseResults.SingleOrDefault(m => m.QualifiedName == target.QualifiedName.QualifiedModuleName);
            if (targetModule == null)
            {
                return null;
            }

            var content = module.Lines[target.Selection.StartLine, 1];

            if (target.DeclarationType == DeclarationType.Parameter)
            {
                var argContext = (VBAParser.ArgContext)target.Context;
                var rewriter = targetModule.GetRewriter();
                rewriter.Replace(argContext.ambiguousIdentifier().Start.TokenIndex, _view.NewName);

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
            return GetReplacementLine(content, newName, target.Selection);
        }*/
    }
}

