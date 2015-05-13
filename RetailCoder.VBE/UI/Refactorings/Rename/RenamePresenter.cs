using System;
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
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Refactorings.Rename
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
            AcquireTarget(_selection);
            if (_view.Target != null)
            {
                _view.ShowDialog();
            }
        }

        public void Show(Declaration target)
        {
            PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;
            _view.ShowDialog();
        }

        private static readonly DeclarationType[] ModuleDeclarationTypes =
            {
                DeclarationType.Class,
                DeclarationType.Module
            };

        private void OnOkButtonClicked(object sender, EventArgs e)
        {
            // must rename usages first; if target is a module or a project,
            // then renaming the declaration first would invalidate the parse results.
            RenameUsages();

            if (ModuleDeclarationTypes.Contains(_view.Target.DeclarationType))
            {
                RenameModule();
            }
            else
            {
                if (_view.Target.DeclarationType == DeclarationType.Project)
                {
                    RenameProject();
                }
                else
                {
                    RenameDeclaration();
                }
            }
        }

        private void RenameModule()
        {
            try
            {
                var module = _view.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                if (module != null)
                {
                    module.Name = _view.NewName;
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

                    var content = module.get_Lines(handler.Selection.StartLine, 1);
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

        private void RenameUsages()
        {
            // todo: refactor

            // rename interface member
            if (_declarations.FindInterfaceMembers().Contains(_view.Target))
            {
                var implementations = _declarations.FindInterfaceImplementationMembers()
                    .Where(m => m.IdentifierName == _view.Target.ComponentName + '_' + _view.Target.IdentifierName);

                foreach (var member in implementations)
                {
                    try
                    {
                        var newMemberName = _view.Target.ComponentName + '_' + _view.NewName;
                        var module = member.Project.VBComponents.Item(member.ComponentName).CodeModule;

                        var content = module.get_Lines(member.Selection.StartLine, 1);
                        var newContent = GetReplacementLine(content, member.IdentifierName, newMemberName);
                        module.ReplaceLine(member.Selection.StartLine, newContent);
                    }
                    catch (COMException)
                    {
                        // gulp
                    }
                }

                return;
            }

            var modules = _view.Target.References.GroupBy(r => r.QualifiedModuleName);
            foreach (var grouping in modules)
            {
                var module = grouping.Key.Component.CodeModule;
                foreach (var line in grouping.GroupBy(reference => reference.Selection.StartLine))
                {
                    var content = module.get_Lines(line.Key, 1);
                    var newContent = GetReplacementLine(content, _view.Target.IdentifierName, _view.NewName);
                    module.ReplaceLine(line.Key, newContent);
                }

                // renaming interface
                if (grouping.Any(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext))
                {
                    var members = _declarations.FindMembers(_view.Target);
                    foreach (var member in members)
                    {
                        var oldMemberName = _view.Target.IdentifierName + '_' + member.IdentifierName;
                        var newMemberName = _view.NewName + '_' + member.IdentifierName;
                        var method = _declarations[oldMemberName].SingleOrDefault(m => m.QualifiedName.QualifiedModuleName == grouping.Key);
                        if (method == null)
                        {
                            continue;
                        }

                        var content = module.get_Lines(method.Selection.StartLine, 1);
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

            var content = module.get_Lines(_view.Target.Selection.StartLine, 1);

            if (target.DeclarationType == DeclarationType.Parameter)
            {
                var argContext = (VBAParser.ArgContext)_view.Target.Context;
                targetModule.Rewriter.Replace(argContext.ambiguousIdentifier().Start.TokenIndex, _view.NewName);

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

                return targetModule.Rewriter.GetText(new Interval(firstTokenIndex, lastTokenIndex));
            }
            else
            {
                return GetReplacementLine(content, target.IdentifierName, newName);
            }
        }

        private static readonly DeclarationType[] ProcedureDeclarationTypes =
            {
                DeclarationType.Procedure,
                DeclarationType.Function,
                DeclarationType.PropertyGet,
                DeclarationType.PropertyLet,
                DeclarationType.PropertySet
            };

        private void AcquireTarget(QualifiedSelection selection)
        {
            var target = _declarations.Items
                .Where(item => !item.IsBuiltIn && item.DeclarationType != DeclarationType.ModuleOption)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item) 
                                      || IsSelectedReference(selection, item));

            PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;

            if (_view.Target == null)
            {
                return;

                // rename the containing procedure:
                _view.Target = _declarations.Items.SingleOrDefault(
                    item => !item.IsBuiltIn 
                            && ProcedureDeclarationTypes.Contains(item.DeclarationType)
                            && item.Context.GetSelection().Contains(selection.Selection));
            }

            if (_view.Target == null)
            {
                return;
                // rename the containing module:
                _view.Target = _declarations.Items.SingleOrDefault(item => 
                    !item.IsBuiltIn
                    && ModuleDeclarationTypes.Contains(item.DeclarationType)
                    && item.QualifiedName.QualifiedModuleName == selection.QualifiedName);
            }
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
