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
    public class RenameRefactoring : IRefactoring
    {
        private readonly IRefactoringPresenterFactory<IRenamePresenter> _factory;
        private readonly IActiveCodePaneEditor _editor;
        private readonly IMessageBox _messageBox;
        private RenameModel _model;

        public RenameRefactoring(IRefactoringPresenterFactory<IRenamePresenter> factory, IActiveCodePaneEditor editor, IMessageBox messageBox)
        {
            _factory = factory;
            _editor = editor;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            _model = presenter.Show();

            if (_model != null && _model.Declarations != null)
            {
                Rename();
            }
        }

        public void Refactor(QualifiedSelection target)
        {
            _editor.SetSelection(target);
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            var presenter = _factory.Create();
            _model = presenter.Show(target);

            if (_model != null && _model.Declarations != null)
            {
                Rename();
            }
        }

        private Declaration FindDeclarationForIdentifier()
        {
            var values = _model.Declarations.Items.Where(item => (item.Scope.Contains(_model.Target.Scope)
                                              || (item.ParentScope == null && _model.Target.ParentScope == null 
                                              || (item.ParentScope != null && _model.Target.ParentScope.Contains(item.ParentScope))))
                                              && _model.NewName == item.IdentifierName).ToList();

            if (values.Any())
            {
                return values.FirstOrDefault();
            }

            foreach (var reference in _model.Target.References)
            {
                var targetReference = reference;
                var potentialDeclarations = _model.Declarations.Items.Where(item => !item.IsBuiltIn
                                                         && item.Project.Equals(targetReference.Declaration.Project)
                                                         && ((item.Context != null
                                                         && item.Context.Start.Line <= targetReference.Selection.StartLine
                                                         && item.Context.Stop.Line >= targetReference.Selection.EndLine)
                                                         || (item.Selection.StartLine <= targetReference.Selection.StartLine
                                                         && item.Selection.EndLine >= targetReference.Selection.EndLine))
                                                         && item.QualifiedName.QualifiedModuleName.ComponentName == targetReference.QualifiedModuleName.ComponentName);

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
                                              && _model.NewName == item.IdentifierName).ToList();

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

        private void Rename()
        {
            var declaration = FindDeclarationForIdentifier();
            if (declaration != null)
            {
                var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _model.NewName, declaration.IdentifierName);
                var rename = _messageBox.Show(message, RubberduckUI.RenameDialog_Caption,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                if (rename == DialogResult.No)
                {
                    return;
                }
            }

            // must rename usages first; if target is a module or a project,
            // then renaming the declaration first would invalidate the parse results.

            if (_model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                // properties can have more than 1 member.
                var members = _model.Declarations[_model.Target.IdentifierName]
                    .Where(item => item.Project == _model.Target.Project
                        && item.ComponentName == _model.Target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property));
                foreach (var member in members)
                {
                    RenameUsages(member);
                }
            }
            else
            {
                RenameUsages(_model.Target);
            }

            if (ModuleDeclarationTypes.Contains(_model.Target.DeclarationType))
            {
                RenameModule();
            }
            else if (_model.Target.DeclarationType == DeclarationType.Project)
            {
                RenameProject();
            }
            else
            {
                RenameDeclaration(_model.Target, _model.NewName);
            }
        }

        private void RenameModule()
        {
            try
            {
                var module = _model.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                if (module != null)
                {
                    if (module.Parent.Type == vbext_ComponentType.vbext_ct_Document)
                    {
                        module.Parent.Properties.Item("_CodeName").Value = _model.NewName;
                    }
                    else if (module.Parent.Type == vbext_ComponentType.vbext_ct_MSForm)
                    {
                        if ((string) module.Parent.Properties.Item("Caption").Value == _model.Target.IdentifierName)
                        {
                            module.Parent.Properties.Item("Caption").Value = _model.NewName;
                        }
                        var codeModule = (CodeModuleClass)module;
                        codeModule.Parent.Name = _model.NewName;
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
                var project = _model.VBE.VBProjects.Cast<VBProject>().FirstOrDefault(p => p.Name == _model.Target.IdentifierName);
                if (project != null)
                {
                    project.Name = _model.NewName;
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

            var module = target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            var newContent = GetReplacementLine(module, target, newName);

            if (target.DeclarationType == DeclarationType.Parameter)
            {
                var argList = (VBAParser.ArgListContext)target.Context.Parent;
                var lineNum = argList.GetSelection().LineCount;

                module.ReplaceLine(argList.Start.Line, newContent);
                module.DeleteLines(argList.Start.Line + 1, lineNum - 1);
            }
            else if (!target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                module.ReplaceLine(target.Selection.StartLine, newContent);
            }
            else
            {
                var members = _model.Declarations[target.IdentifierName]
                    .Where(item => item.Project == target.Project
                        && item.ComponentName == target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property));

                foreach (var member in members)
                {
                    newContent = GetReplacementLine(module, member, newName);
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
                    var newMemberName = handler.IdentifierName.Replace(control.Name + '_', _model.NewName + '_');
                    var module = handler.Project.VBComponents.Item(handler.ComponentName).CodeModule;

                    var content = module.Lines[handler.Selection.StartLine, 1];
                    var newContent = GetReplacementLine(content, newMemberName, handler.Selection);
                    module.ReplaceLine(handler.Selection.StartLine, newContent);
                }

                control.Name = _model.NewName;
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
                            newContent = GetReplacementLine(content, _model.NewName,
                                reference.Selection);
                        }
                        else
                        {
                            newContent = GetReplacementLine(content,
                                interfaceName + "_" + _model.NewName,
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
                        var newMemberName = _model.NewName + '_' + member.IdentifierName;
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
                rewriter.Replace(argContext.ambiguousIdentifier().Start.TokenIndex, _model.NewName);

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
