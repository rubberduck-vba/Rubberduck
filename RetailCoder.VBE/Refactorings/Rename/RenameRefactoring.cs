﻿using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Microsoft.CSharp.RuntimeBinder;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringPresenterFactory<IRenamePresenter> _factory;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private RenameModel _model;

        public RenameRefactoring(IVBE vbe, IRefactoringPresenterFactory<IRenamePresenter> factory, IMessageBox messageBox, RubberduckParserState state)
        {
            _vbe = vbe;
            _factory = factory;
            _messageBox = messageBox;
            _state = state;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            _model = presenter.Show();

            QualifiedSelection? oldSelection = null;
            var pane = _vbe.ActiveCodePane;
            var module = pane.CodeModule;
            if (!pane.IsWrappingNullReference)
            {
                oldSelection = module.GetQualifiedSelection();
            }

            if (_model != null && _model.Declarations != null)
            {
                Rename();
            }

            if (oldSelection.HasValue)
            {
                pane.Selection = oldSelection.Value.Selection;
            }
        }

        public void Refactor(QualifiedSelection target)
        {
            var pane = _vbe.ActiveCodePane;
            if (pane.IsWrappingNullReference)
            {
                return;
            }
            pane.Selection = target.Selection;
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            if (target.IsBuiltIn) { return; }

            var presenter = _factory.Create();
            _model = presenter.Show(target);

            var oldSelection = Selection.Home;
            var pane = _vbe.ActiveCodePane;
            if (!pane.IsWrappingNullReference)
            {
                oldSelection = pane.Selection;
            }

            if (_model != null && _model.Declarations != null)
            {
                Rename();
            }

            if (!pane.IsWrappingNullReference)
            {
                pane.Selection = oldSelection;
            }
        }
        private static readonly DeclarationType[] ModuleDeclarationTypes =
        {
            DeclarationType.ClassModule,
            DeclarationType.ProceduralModule
        };

        private void Rename()
        {
            var declaration = _state.DeclarationFinder.GetDeclarationsAccessibleToScope(_model.Target, _model.Declarations)
                .Where(d => d.IdentifierName.Equals(_model.NewName, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (declaration != null)
            {
                var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _model.NewName,
                    declaration.IdentifierName);
                var rename = _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation);

                if (rename == DialogResult.No)
                {
                    return;
                }
            }
            else if(_model.Target == null)
            {
                return;
            }

            // must rename usages first; if target is a module or a project,
            // then renaming the declaration first would invalidate the parse results.
            Debug.Assert(_model.Target != null);

            if (_model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                // properties can have more than 1 member.
                var members = _model.Declarations.Named(_model.Target.IdentifierName)
                    .Where(item => item.ProjectId == _model.Target.ProjectId
                        && item.ComponentName == _model.Target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property));
                foreach (var member in members)
                {
                    RenameUsages(member);
                }
            }
            else if (_model.Target.DeclarationType == DeclarationType.Parameter && _model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var getter = _model.Target.DeclarationType == DeclarationType.PropertyGet
                    ? _model.Target
                    : GetProperty(_model.Target.ParentDeclaration, DeclarationType.PropertyGet);

                var letter = _model.Target.DeclarationType == DeclarationType.PropertyLet
                    ? _model.Target
                    : GetProperty(_model.Target.ParentDeclaration, DeclarationType.PropertyLet);

                var setter = _model.Target.DeclarationType == DeclarationType.PropertySet
                    ? _model.Target
                    : GetProperty(_model.Target.ParentDeclaration, DeclarationType.PropertySet);

                var properties = new[] {getter, letter, setter};

                var parameters = _model.Declarations.Where(d =>
                    d.DeclarationType == DeclarationType.Parameter &&
                    properties.Contains(d.ParentDeclaration) &&
                    d.IdentifierName == _model.Target.IdentifierName);

                foreach (var param in parameters)
                {
                    RenameUsages(param);
                    RenameDeclaration(param, _model.NewName);
                }
            }
            else
            {
                RenameUsages(_model.Target);
            }

            if (ModuleDeclarationTypes.Contains(_model.Target.DeclarationType))
            {
                RenameModule();
                return; // renaming a component automatically triggers a reparse
            }
            else if (_model.Target.DeclarationType == DeclarationType.Project)
            {
                RenameProject();
                return; // renaming a project automatically triggers a reparse
            }
            else
            {
                // we handled properties above
                if (!_model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
                {
                    RenameDeclaration(_model.Target, _model.NewName);
                }
            }

            _state.OnParseRequested(this);
        }
        
        private Declaration GetProperty(Declaration declaration, DeclarationType declarationType)
        {
            return _model.Declarations.FirstOrDefault(item => item.Scope == declaration.Scope &&
                              item.IdentifierName == declaration.IdentifierName &&
                              item.DeclarationType == declarationType);
        }

        private void RenameModule()
        {
            try
            {
                var component = _model.Target.QualifiedName.QualifiedModuleName.Component;
                var module = component.CodeModule;
                {
                    if (module.IsWrappingNullReference)
                    {
                        return;
                    }

                    if (component.Type == ComponentType.Document)
                    {
                        var properties = component.Properties;
                        var property = properties["_CodeName"];
                        {
                            property.Value = _model.NewName;
                        }
                    }
                    else if (component.Type == ComponentType.UserForm)
                    {
                        var properties = component.Properties;
                        var property = properties["Caption"];
                        {
                            if ((string)property.Value == _model.Target.IdentifierName)
                            {
                                property.Value = _model.NewName;
                            }
                            component.Name = _model.NewName;
                        }
                    }
                    else
                    {
                        module.Name = _model.NewName;
                    }
                }
            }
            catch (COMException)
            {
                _messageBox.Show(RubberduckUI.RenameDialog_ModuleRenameError, RubberduckUI.RenameDialog_Caption);
            }
        }

        private void RenameProject()
        {
            try
            {
                var projects = _vbe.VBProjects;
                var project = projects.SingleOrDefault(p => p.HelpFile == _model.Target.ProjectId);
                {
                    if (project != null)
                    {
                        project.Name = _model.NewName;
                    }
                }
            }
            catch (COMException)
            {
                _messageBox.Show(RubberduckUI.RenameDialog_ProjectRenameError, RubberduckUI.RenameDialog_Caption);
            }
        }

        private void RenameDeclaration(Declaration target, string newName)
        {
            if (target.DeclarationType == DeclarationType.Control)
            {
                RenameControl();
                return;
            }

            var component = target.QualifiedName.QualifiedModuleName.Component;
            var module = component.CodeModule;
            {
                var newContent = GetReplacementLine(module, target, newName);

                if (target.DeclarationType == DeclarationType.Parameter)
                {
                    var argList = (VBAParser.ArgListContext)target.Context.Parent;
                    var lineNum = argList.GetSelection().LineCount;

                    // delete excess lines to prevent removing our own changes
                    module.DeleteLines(argList.Start.Line + 1, lineNum - 1);
                    module.ReplaceLine(argList.Start.Line, newContent);
                    
                }
                else if (!target.DeclarationType.HasFlag(DeclarationType.Property))
                {
                    module.ReplaceLine(target.Selection.StartLine, newContent);
                }
                else
                {
                    var members = _model.Declarations.Named(target.IdentifierName)
                        .Where(item => item.ProjectId == target.ProjectId
                            && item.ComponentName == target.ComponentName
                            && item.DeclarationType.HasFlag(DeclarationType.Property));

                    foreach (var member in members)
                    {
                        newContent = GetReplacementLine(module, member, newName);
                        module.ReplaceLine(member.Selection.StartLine, newContent);
                    }
                }
            }
        }

        private void RenameControl()
        {
            try
            {
                var module = _model.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                var component = module.Parent;
                var control = component.Controls.SingleOrDefault(item => item.Name == _model.Target.IdentifierName);
                {
                    if (control == null)
                    {
                        return;
                    }

                    foreach (var handler in _model.Declarations.FindEventHandlers(_model.Target).OrderByDescending(h => h.Selection.StartColumn))
                    {
                        var newMemberName = handler.IdentifierName.Replace(control.Name + '_', _model.NewName + '_');
                        var project = handler.Project;
                        var components = project.VBComponents;
                        var refComponent = components[handler.ComponentName];
                        var refModule = refComponent.CodeModule;
                        {
                            var content = refModule.GetLines(handler.Selection.StartLine, 1);
                            var newContent = GetReplacementLine(content, newMemberName, handler.Selection);
                            refModule.ReplaceLine(handler.Selection.StartLine, newContent);
                        }
                    }

                    control.Name = _model.NewName;
                }
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


            if (target.DeclarationType == DeclarationType.Event)
            {
                var handlers = _model.Declarations.FindHandlersForEvent(target);
                foreach (var handler in handlers)
                {
                    RenameDeclaration(handler.Item2, handler.Item1.IdentifierName + '_' + _model.NewName);
                }
            }

            // rename interface member
            if (_model.Declarations.FindInterfaceMembers().Contains(target))
            {
                var implementations = _model.Declarations.FindInterfaceImplementationMembers()
                    .Where(m => m.IdentifierName == target.ComponentName + '_' + target.IdentifierName);

                foreach (var member in implementations.OrderByDescending(m => m.Selection.StartColumn))
                {
                    try
                    {
                        var newMemberName = target.ComponentName + '_' + _model.NewName;
                        var project = member.Project;
                        var components = project.VBComponents;
                        var component = components[member.ComponentName];
                        var module = component.CodeModule;
                        {
                            var content = module.GetLines(member.Selection.StartLine, 1);
                            var newContent = GetReplacementLine(content, newMemberName, member.Selection);
                            module.ReplaceLine(member.Selection.StartLine, newContent);
                            RenameUsages(member, target.ComponentName);
                        }
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
                {
                    foreach (var line in grouping.GroupBy(reference => reference.Selection.StartLine))
                    {
                        var lastSelection = Selection.Empty;
                        foreach (var reference in line.OrderByDescending(l => l.Selection.StartColumn))
                        {
                            if (reference.Selection == lastSelection)
                            {
                                continue;
                            }

                            var content = module.GetLines(line.Key, 1);
                            string newContent;

                            if (interfaceName == null)
                            {
                                newContent = GetReplacementLine(content, _model.NewName, reference.Selection);
                            }
                            else
                            {
                                newContent = GetReplacementLine(content, interfaceName + "_" + _model.NewName, reference.Selection);
                            }

                            module.ReplaceLine(line.Key, newContent);
                            lastSelection = reference.Selection;
                        }
                    }

                    // renaming interface
                    if (grouping.Any(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext))
                    {
                        var members = _model.Declarations.InScope(target).OrderByDescending(m => m.Selection.StartColumn);
                        foreach (var member in members)
                        {
                            var oldMemberName = target.IdentifierName + '_' + member.IdentifierName;
                            var newMemberName = _model.NewName + '_' + member.IdentifierName;
                            var method = _model.Declarations.Named(oldMemberName).SingleOrDefault(m => m.QualifiedName.QualifiedModuleName == grouping.Key);
                            if (method == null)
                            {
                                continue;
                            }

                            var content = module.GetLines(method.Selection.StartLine, 1);
                            var newContent = GetReplacementLine(content, newMemberName, member.Selection);
                            module.ReplaceLine(method.Selection.StartLine, newContent);
                        }
                    }
                }
            }
        }

        private string GetReplacementLine(string content, string newName, Selection selection)
        {
            var contentWithoutOldName = content.Remove(selection.StartColumn - 1, selection.EndColumn - selection.StartColumn);
            return contentWithoutOldName.Insert(selection.StartColumn - 1, newName);
        }

        private string GetReplacementLine(ICodeModule module, Declaration target, string newName)
        {
            var content = module.GetLines(target.Selection.StartLine, 1);

            if (target.DeclarationType == DeclarationType.Parameter)
            {
                var argContext = (VBAParser.ArgContext)target.Context;
                var rewriter = _model.State.GetRewriter(target.QualifiedName.QualifiedModuleName.Component);
                rewriter.Replace(argContext.unrestrictedIdentifier().Start.TokenIndex, _model.NewName);

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
        }
    }
}
