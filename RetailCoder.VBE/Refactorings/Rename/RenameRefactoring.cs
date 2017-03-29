using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Antlr4.Runtime;
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
    internal interface IRenameRefactoringHandler
    {
        void Rename();
        string ErrorMessage { get; }
    }

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

            QualifiedSelection? qOldSelection = null;
            var pane = _vbe.ActiveCodePane;
            if (!pane.IsWrappingNullReference)
            {
                qOldSelection = pane.CodeModule.GetQualifiedSelection();
            }

            Rename();

            if (qOldSelection.HasValue)
            {
                pane.Selection = qOldSelection.Value.Selection;
            }
        }

        //The User has selected something in a code pane
        public void Refactor(QualifiedSelection qSelection)
        {
            var pane = _vbe.ActiveCodePane;
            if (pane.IsWrappingNullReference)
            {
                return;
            }
            pane.Selection = qSelection.Selection;
            Refactor();
        }

        //The refactor starts from a Declaration only...not sure where the user selection is(?)
        public void Refactor(Declaration target)
        {
            if (!target.IsUserDefined)
            {
                //Todo: Add format string to UI resources
                var format = "Unable to rename Built-In declarations like '{0}'";
                var message = string.Format(format, target.IdentifierName);
                _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            var presenter = _factory.Create();
            _model = presenter.Show(target);

            var uqOldSelection = Selection.Home;
            var pane = _vbe.ActiveCodePane;
            if (!pane.IsWrappingNullReference)
            {
                uqOldSelection = pane.Selection;
            }

            Rename();

            if (!pane.IsWrappingNullReference)
            {
                pane.Selection = uqOldSelection;
            }
        }

        private DialogResult ObtainApprovalToUseConflictingName(Declaration declaration)
        {
            var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _model.NewName,
                declaration);
            return _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.YesNo,
                MessageBoxIcon.Exclamation);
        }

        private bool UserCancelsRename()
        {
            var declarations = _state.DeclarationFinder.GetDeclarationsWithIdentifiersToAvoid(_model.Target)
                .Where(d => d.IdentifierName.Equals(_model.NewName, StringComparison.InvariantCultureIgnoreCase));//.FirstOrDefault();

            if (declarations.Any())
            {
                var rename = ObtainApprovalToUseConflictingName(declarations.FirstOrDefault());

                if (rename == DialogResult.No)
                {
                    return true;
                }
            }
            return false;
        }

        private void Rename()
        {
            if (_model == null || _model.Declarations == null || _model.Target == null) { return; }

            if (UserCancelsRename()) { return; }

            //todo: remove original code option and internal class
            if (UseOriginalImplementation()) { return; }

            var handler = GetHandler(_model);

            try
            {
                handler.Rename();
            }
            catch (Exception)
            {
                _messageBox.Show(handler.ErrorMessage, RubberduckUI.RenameDialog_Caption);
            }
        }

        private static IRenameRefactoringHandler GetHandler(RenameModel model)
        {
            IRenameRefactoringHandler handler = null;
            if (model.Target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                handler = new RenamePropertyHandler(model);
            }
            else if (model.Target.DeclarationType == DeclarationType.Parameter
                        && model.Target.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                handler = new RenamePropertyParameterHandler(model);
            }
            else if (model.Target.DeclarationType.HasFlag(DeclarationType.Module))
            {
                handler = new RenameModuleHandler(model);
            }
            else if (model.Target.DeclarationType == DeclarationType.Project)
            {
                handler = new RenameProjectHandler(model);
            }
            else if (model.Declarations.FindInterfaceMembers().Contains(model.Target))
            {
                handler = new RenameInterfaceHandler(model);
            }
            else if (model.Target.DeclarationType == DeclarationType.Event || IsEvent(model))
            {
                handler = new RenameEventHandler(model);
            }
            else if (model.Target.DeclarationType == DeclarationType.Control || IsControl(model))
            {
                handler = new RenameControlHandler(model);
            }
            else
            {
                handler = new RenameDefaultHandler(model);
            }
            return handler;
        }

        private static bool IsControl(RenameModel model)
        {
            var declaration = GetDeclarationOfType(model, DeclarationType.Control, false);
            return (declaration != null);
        }

        private static bool IsEvent(RenameModel model)
        {
            var declaration = GetDeclarationOfType(model, DeclarationType.Event, true);
            return (declaration != null);
        }

        private static Declaration GetDeclarationOfType(RenameModel model, DeclarationType goalType, bool identifierIsOnRHS)
        {
            var declarationName = model.Target.IdentifierName;
            try
            {
                if (model.Target.IdentifierName.Contains("_"))
                {
                    if (identifierIsOnRHS)
                    {
                        //Events
                        declarationName = model.Target.IdentifierName.Remove(0,model.Target.IdentifierName.IndexOf("_")+1);
                    }
                    else
                    {
                        //Controls
                        declarationName = model.Target.IdentifierName.Remove(model.Target.IdentifierName.IndexOf("_"));
                    }
                }
                var declaration = model.Declarations.Single(item =>
                    item.DeclarationType == goalType
                    && (declarationName.Equals(item.IdentifierName)));
                return declaration;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void RenameUsages(RenameModel model, Declaration target)
        {
            RenameUsages(model, target, model.NewName);
        }

        private static void RenameUsages(RenameModel model, Declaration target, string newName)
        {
            var modules = target.References.GroupBy(r => r.QualifiedModuleName);
            foreach (var grouping in modules)
            {
                var rewriter = model.GetRewriter(grouping.Key);
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
                            var newContent = reference.Context.GetText().Replace(reference.IdentifierName, newName);
                            rewriter.Replace(reference.Context, newContent);
                            lastSelection = reference.Selection;
                        }
                    }
                }
            }
        }

        private static void RenameDeclaration(RenameModel model, Declaration target)
        {
            RenameDeclaration(model, target, model.NewName);
        }

        private static void RenameDeclaration(RenameModel model, Declaration target, string newName)
        {

            var component = target.QualifiedName.QualifiedModuleName.Component;
            var rewriter = model.GetRewriter(target);
            var module = component.CodeModule;
            {
                if (target.DeclarationType == DeclarationType.Parameter)
                {
                    var argList = (VBAParser.ArgListContext)target.Context.Parent;
                    var lineNum = argList.GetSelection().LineCount;

                    var newContent = target.Context.GetText().Replace(target.IdentifierName, newName);
                    rewriter.Replace(target, newContent );
                }
                else if (!target.DeclarationType.HasFlag(DeclarationType.Property))
                {
                    var newContent = target.Context.GetText().Replace(target.IdentifierName, newName);
                    rewriter.Replace(target.Context, newContent);
                }
                else
                {
                    var members = model.Declarations.Named(target.IdentifierName)
                        .Where(item => item.ProjectId == target.ProjectId
                            && item.ComponentName == target.ComponentName
                            && item.DeclarationType.HasFlag(DeclarationType.Property));

                    foreach (var member in members)
                    {
                        var newContent = member.Context.GetText().Replace(member.IdentifierName, newName);
                        rewriter.Replace(member.Context, newContent);
                    }
                }
            }
        }

        //todo: decide if the following internal classes should become full-fledged independent classes

        internal class RenameDefaultHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenameDefaultHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_DefaultRenameError"; } }

            public void Rename()
            {
                RenameUsages(_model, _model.Target);
                RewriteContent(_model);

                RenameDeclaration(_model, _model.Target);
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenamePropertyHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenamePropertyHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_PropertyRenameError"; } }


            public void Rename()
            {
                // properties can have more than 1 member.
                var members = _model.Declarations.Named(_model.Target.IdentifierName)
                    .Where(item => item.ProjectId == _model.Target.ProjectId
                        && item.ComponentName == _model.Target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property)).ToList();

                members.ForEach(member => RenameUsages(_model, member));
                RewriteContent(_model);

                RenameDeclaration(_model, _model.Target);
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenamePropertyParameterHandler : IRenameRefactoringHandler
        {

            private RenameModel _model;

            public RenamePropertyParameterHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_PropertyParameterRenameError"; } }

            public void Rename()
            {
                var parameters = _model.Declarations.Where(d =>
                    d.DeclarationType == DeclarationType.Parameter
                    && d.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property)
                    && d.IdentifierName == _model.Target.IdentifierName).ToList();

                parameters.ForEach(param => RenameUsages(_model, param));
                RewriteContent(_model);

                parameters.ForEach(param => RenameDeclaration(_model, param));
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenameInterfaceHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenameInterfaceHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_InterfaceRenameError"; } }

            public void Rename()
            {
                RenameUsages(_model, _model.Target);
                RewriteContent(_model);

                var implementations = _model.Declarations.FindInterfaceImplementationMembers()
                    .Where(m => m.IdentifierName == _model.Target.ComponentName + '_' + _model.Target.IdentifierName)
                        .OrderByDescending(m => m.Selection.StartColumn).ToList();

                var newMemberName = _model.Target.ComponentName + '_' + _model.NewName;
                implementations.ForEach(imp => RenameDeclaration(_model, imp, newMemberName));
                RewriteContent(_model);

                RenameDeclaration(_model, _model.Target);
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenameEventHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenameEventHandler(RenameModel model)
            {
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_EventRenameError"; } }

            public void Rename()
            {
                var eventDeclaration = GetDeclarationOfType(_model, DeclarationType.Event, true);
                if (null == eventDeclaration) { return; }

                var newEventName = _model.NewName;
                if (newEventName.Contains("_"))
                {
                    newEventName = _model.NewName.Remove(0,_model.NewName.IndexOf("_")+1);
                }

                RenameUsages(_model, eventDeclaration, newEventName);
                RewriteContent(_model);

                var handlers = _model.Declarations.FindHandlersForEvent(eventDeclaration).ToList();
                handlers.ForEach(handler => RenameDeclaration(_model, handler.Item2, handler.Item1.IdentifierName + '_' + newEventName));
                RewriteContent(_model);

                RenameDeclaration(_model, eventDeclaration, newEventName);
                RewriteContent(_model);

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenameControlHandler : IRenameRefactoringHandler
        {
            private Declaration _target;
            private RenameModel _model;

            public RenameControlHandler(RenameModel model)
            {
                _target = model.Target;
                _model = model;
            }

            //todo: Add error message string to UI Resources
            public string ErrorMessage { get { return "RenameDialog_ControlRenameError"; } }

            public void Rename()
            {
                var controlDeclaration = GetDeclarationOfType(_model, DeclarationType.Control, false);
                if (null == controlDeclaration) { return; }

                var module = _model.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                var component = module.Parent;
                var control = component.Controls.SingleOrDefault(item => item.Name == controlDeclaration.IdentifierName);

                if (control == null) { return; }

                var newControlName = _model.NewName;
                if (newControlName.Contains("_"))
                {
                    newControlName = _model.NewName.Remove(_model.NewName.IndexOf("_"));
                }

                //User attempted to rename an event sub signature e.g. btn1_Click() to btn1_ClickAgain()
                if (newControlName.Equals(controlDeclaration.IdentifierName))
                {
                    //nothing to change
                    return;
                }

                RenameUsages(_model, controlDeclaration, newControlName);
                RewriteContent(_model);

                var handlers = _model.Declarations.FindEventHandlers(controlDeclaration).OrderByDescending(h => h.Selection.StartColumn).ToList();

                foreach (var handler in handlers)
                {
                    var newMemberName = handler.IdentifierName.Replace(control.Name + '_', newControlName + '_');
                    RenameUsages(_model, handler, newMemberName);
                }
                RewriteContent(_model);

                foreach (var handler in handlers)
                {
                    var newMemberName = handler.IdentifierName.Replace(control.Name + '_', newControlName + '_');
                    RenameDeclaration(_model, handler, newMemberName);
                }
                RewriteContent(_model);

                control.Name = newControlName;

                _model.State.OnParseRequested(this);
            }
        }

        internal class RenameProjectHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;
            public RenameProjectHandler(RenameModel model)
            {
                _model = model;
            }

            public string ErrorMessage {  get { return RubberduckUI.RenameDialog_ProjectRenameError;  } }

            public void Rename()
            {
                var projects = _model.VBE.VBProjects;
                var project = projects.SingleOrDefault(p => p.HelpFile == _model.Target.ProjectId);
                {
                    if (project != null)
                    {
                        project.Name = _model.NewName;
                    }
                }
            }
        }

        internal class RenameModuleHandler : IRenameRefactoringHandler
        {
            private RenameModel _model;

            public RenameModuleHandler(RenameModel model)
            {
                _model = model;
            }

            public string ErrorMessage { get { return RubberduckUI.RenameDialog_ModuleRenameError; } }

            public void Rename()
            {
                RenameUsages(_model, _model.Target);
                RewriteContent(_model);

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
        }

        private static void RewriteContent(RenameModel model)
        {
            model.Rewrite();
            model.ClearRewriters();
        }

/***************************************************************************************/
        private bool UseOriginalImplementation()
        {
            //todo: remove original code option and internal class
            bool executeUsingOriginalImplementation = false;
            if (executeUsingOriginalImplementation)
            {
                var legacy = new OriginalRenameCode(_vbe, _factory, _messageBox, _state);
                legacy.Rename();
            }
            return executeUsingOriginalImplementation;
        }
/***************************************************************************************/

        internal class OriginalRenameCode : IRenameRefactoringHandler
        {
            private readonly IVBE _vbe;
            private readonly IRefactoringPresenterFactory<IRenamePresenter> _factory;
            private readonly IMessageBox _messageBox;
            private readonly RubberduckParserState _state;
            private RenameModel _model;

            public OriginalRenameCode(IVBE vbe, IRefactoringPresenterFactory<IRenamePresenter> factory, IMessageBox messageBox, RubberduckParserState state)
            {
                _vbe = vbe;
                _factory = factory;
                _messageBox = messageBox;
                _state = state;
            }

            public bool RequestReparseAfterRename { get { return false; } }

            public string ErrorMessage { get { return ""; } }

            private static readonly DeclarationType[] ModuleDeclarationTypes =
            {
                DeclarationType.ClassModule,
                DeclarationType.ProceduralModule
            };

            public void Rename()
            {
                var presenter = _factory.Create();
                _model = presenter.Show();

                var declaration = _state.DeclarationFinder
                    .GetDeclarationsWithIdentifiersToAvoid(_model.Target)
                    .FirstOrDefault(d => d.IdentifierName.Equals(_model.NewName, StringComparison.InvariantCultureIgnoreCase));
                if (declaration != null)
                {
                    var message = string.Format(RubberduckUI.RenameDialog_ConflictingNames, _model.NewName,
                        declaration);
                    var rename = _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.YesNo,
                        MessageBoxIcon.Exclamation);

                    if (rename == DialogResult.No)
                    {
                        return;
                    }
                }
                else if (_model.Target == null)
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

                    var properties = new[] { getter, letter, setter };

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
                    var rewriter = _model.State.GetRewriter(target.QualifiedName.QualifiedModuleName.Component);

                    var identifier = ((VBAParser.ArgContext)target.Context).unrestrictedIdentifier();
                    rewriter.Replace(identifier, _model.NewName);

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

                    return rewriter.GetText(firstTokenIndex, lastTokenIndex);
                }
                return GetReplacementLine(content, newName, target.Selection);
            }

    }
    }
}
