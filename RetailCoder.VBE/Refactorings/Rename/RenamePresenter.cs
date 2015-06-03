using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
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
        private readonly VBE _vbe;
        private readonly IRenameView _view;
        private readonly Declarations _declarations;
        private readonly QualifiedSelection _selection;
        private readonly VBProjectParseResult _parseResult;

        public RenamePresenter(VBE vbe, IRenameView view, VBProjectParseResult parseResult, QualifiedSelection selection)
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
            Declaration target;
            AcquireTarget(out target, _selection);
            if (target != null)
            {
                _view.Target = target;
                _view.ShowDialog();
            }
        }

        public void Show(Declaration target)
        {
            PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;
            _view.ShowDialog();
        }

        private Declaration AmbiguousId()
        {
            var values = _declarations.Items.Where(item => (item.Scope.Contains(_view.Target.Scope)
                                              || _view.Target.ParentScope.Contains(item.ParentScope))
                                              && _view.NewName == item.IdentifierName);

            if (values.Any())
            {
                return values.FirstOrDefault();
            }

            foreach (var reference in _view.Target.References)
            {
                var potentialDeclarations = _declarations.Items.Where(item => !item.IsBuiltIn
                                                         && item.Project.Equals(reference.Declaration.Project)
                                                         && ((item.Context != null
                                                         && item.Context.Start.Line <= reference.Selection.StartLine
                                                         && item.Context.Stop.Line >= reference.Selection.EndLine)
                                                         || (item.Selection.StartLine <= reference.Selection.StartLine
                                                         && item.Selection.EndLine >= reference.Selection.EndLine)));

                var currentStartLine = 0;
                var currentEndLine = int.MaxValue;
                var currentStartColumn = 0;
                var currentEndColumn = int.MaxValue;

                Declaration target = null;

                foreach (var item in potentialDeclarations)
                {
                    if (currentStartLine <= item.Selection.StartLine && currentEndLine >= item.Selection.EndLine)
                    {
                        if (!(item.Selection.StartLine == reference.Selection.StartLine &&
                              (item.Selection.StartColumn > reference.Selection.StartColumn ||
                               currentStartColumn > item.Selection.StartColumn) ||
                              item.Selection.EndLine == reference.Selection.EndLine &&
                              (item.Selection.EndColumn < reference.Selection.EndColumn ||
                               currentEndColumn < item.Selection.EndColumn)))
                        {
                            currentStartLine = item.Selection.StartLine;
                            currentEndLine = item.Selection.EndLine;
                            currentStartColumn = item.Selection.StartColumn;
                            currentEndColumn = item.Selection.EndColumn;

                            target = item;
                        }
                    }
                }

                if (target == null) { return null; }

                values = _declarations.Items.Where(item => (item.Scope.Contains(target.Scope)
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
            };

        private void OnOkButtonClicked(object sender, EventArgs e)
        {
            var ambiguousId = AmbiguousId();
            if (ambiguousId != null)
            {
                var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _view.NewName,
                    ambiguousId.IdentifierName);
                var rename = MessageBox.Show(message, RubberduckUI.RenameDialog_Caption,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                if (rename == DialogResult.No)
                {
                    return;
                }
            }

            // must rename usages first; if target is a module or a project,
            // then renaming the declaration first would invalidate the parse results.
            RenameUsages(_view.Target);

            if (ModuleDeclarationTypes.Contains(_view.Target.DeclarationType))
            {
                RenameModule();
            }
            else if (_view.Target.DeclarationType == DeclarationType.Project)
            {
                RenameProject();
            }
            else
            {
                RenameDeclaration();
            }
        }

        private void RenameModule()
        {
            try
            {
                var module = _view.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
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
                var project = _vbe.VBProjects.Cast<VBProject>().FirstOrDefault(p => p.Name == _view.Target.IdentifierName);
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
            if (_view.Target.DeclarationType == DeclarationType.Control)
            {
                RenameControl();
                return;
            }

            var module = _view.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            var newContent = GetReplacementLine(module, _view.Target, _view.NewName);

            if (_view.Target.DeclarationType == DeclarationType.Parameter)
            {
                var argList = (VBAParser.ArgListContext)_view.Target.Context.Parent;
                var lineNum = argList.GetSelection().LineCount;

                module.ReplaceLine(argList.Start.Line, newContent);
                module.DeleteLines(argList.Start.Line + 1, lineNum - 1);
            }
            else
            {
                module.ReplaceLine(_view.Target.Selection.StartLine, newContent);
            }
        }

        private void RenameControl()
        {
            try
            {
                var form = _view.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                var control = ((dynamic) form.Parent.Designer).Controls(_view.Target.IdentifierName);

                foreach (var handler in _declarations.FindEventHandlers(_view.Target))
                {
                    var newMemberName = handler.IdentifierName.Replace(control.Name + '_', _view.NewName + '_');
                    var module = handler.Project.VBComponents.Item(handler.ComponentName).CodeModule;

                    var content = module.Lines[handler.Selection.StartLine, 1];
                    var newContent = GetReplacementLine(content, handler.IdentifierName, newMemberName);
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
            if (_declarations.FindInterfaceMembers().Contains(target))
            {
                var implementations = _declarations.FindInterfaceImplementationMembers()
                    .Where(m => m.IdentifierName == target.ComponentName + '_' + target.IdentifierName);

                foreach (var member in implementations)
                {
                    try
                    {
                        var newMemberName = target.ComponentName + '_' + _view.NewName;
                        var module = member.Project.VBComponents.Item(member.ComponentName).CodeModule;

                        var content = module.Lines[member.Selection.StartLine, 1];
                        var newContent = GetReplacementLine(content, member.IdentifierName, newMemberName);
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
                    var content = module.Lines[line.Key, 1];
                    string newContent;

                    if (interfaceName == null)
                    {
                        newContent = GetReplacementLine(content, target.IdentifierName, _view.NewName);
                    }
                    else
                    {
                        newContent = GetReplacementLine(content, target.IdentifierName, interfaceName + "_" + _view.NewName);
                    }

                    module.ReplaceLine(line.Key, newContent);
                }

                // renaming interface
                if (grouping.Any(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext))
                {
                    var members = _declarations.FindMembers(target);
                    foreach (var member in members)
                    {
                        var oldMemberName = target.IdentifierName + '_' + member.IdentifierName;
                        var newMemberName = _view.NewName + '_' + member.IdentifierName;
                        var method = _declarations[oldMemberName].SingleOrDefault(m => m.QualifiedName.QualifiedModuleName == grouping.Key);
                        if (method == null)
                        {
                            continue;
                        }

                        var content = module.Lines[method.Selection.StartLine, 1];
                        var newContent = GetReplacementLine(content, oldMemberName, newMemberName);
                        module.ReplaceLine(method.Selection.StartLine, newContent);
                    }
                }
            }
        }

        private string GetReplacementLine(string content, string target, string newName)
        {
            // until we figure out how to replace actual tokens,
            // this is going to have to be done the ugly way...
            return Regex.Replace(content, "\\b" + target + "\\b", newName);
        }

        private string GetReplacementLine(CodeModule module, Declaration target, string newName)
        {
            var targetModule = _parseResult.ComponentParseResults.SingleOrDefault(m => m.QualifiedName == _view.Target.QualifiedName.QualifiedModuleName);
            if (targetModule == null)
            {
                return null;
            }

            var content = module.Lines[_view.Target.Selection.StartLine, 1];

            if (target.DeclarationType == DeclarationType.Parameter)
            {
                var argContext = (VBAParser.ArgContext)_view.Target.Context;
                var rewriter = targetModule.GetRewriter();
                rewriter.Replace(argContext.ambiguousIdentifier().Start.TokenIndex, _view.NewName);

                // Target.Context is an ArgContext, its parent is an ArgsListContext;
                // the ArgsListContext's parent is the procedure context and it includes the body.
                var context = (ParserRuleContext) _view.Target.Context.Parent.Parent;
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
            return GetReplacementLine(content, target.IdentifierName, newName);
        }

        private static readonly DeclarationType[] ProcedureDeclarationTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private void AcquireTarget(out Declaration target, QualifiedSelection selection)
        {
            target = _declarations.Items
                .Where(item => !item.IsBuiltIn && item.DeclarationType != DeclarationType.ModuleOption)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item) 
                                      || IsSelectedReference(selection, item));

            PromptIfTargetImplementsInterface(ref target);

            /*if (target == null)
            {
                return;

                // rename the containing procedure:
                _view.Target = _declarations.Items.SingleOrDefault(
                    item => !item.IsBuiltIn 
                            && ProcedureDeclarationTypes.Contains(item.DeclarationType)
                            && item.Context.GetSelection().Contains(selection.Selection));
            }

            if (target == null)
            {
                return;
                // rename the containing module:
                _view.Target = _declarations.Items.SingleOrDefault(item => 
                    !item.IsBuiltIn
                    && ModuleDeclarationTypes.Contains(item.DeclarationType)
                    && item.QualifiedName.QualifiedModuleName == selection.QualifiedName);
            }*/
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
                r.QualifiedModuleName.Project == selection.QualifiedName.Project
                && r.QualifiedModuleName.ComponentName == selection.QualifiedName.ComponentName
                && r.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName.Project == selection.QualifiedName.Project
                   && declaration.QualifiedName.QualifiedModuleName.ComponentName == selection.QualifiedName.ComponentName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
        }
    }
}
