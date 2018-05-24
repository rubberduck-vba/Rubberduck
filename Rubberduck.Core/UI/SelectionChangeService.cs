using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI
{
    public interface IAutoComplete
    {
        string InputToken { get; }
        string OutputToken { get; }
        string Execute(TypingCodeEventArgs e);
        bool IsEnabled { get; }
    }

    public abstract class AutoCompleteBase : IAutoComplete
    {
        public bool IsEnabled => true;
        public abstract string InputToken { get; }
        public abstract string OutputToken { get; }

        public virtual string Execute(TypingCodeEventArgs e)
        {
            var selection = e.CodePane.Selection;
            if (selection.StartColumn < 2) { return null; }

            if (!e.IsCommitted && e.Code.Substring(selection.StartColumn - 2, 1) == InputToken)
            {
                using (var module = e.CodePane.CodeModule)
                {
                    var replacement = e.Code.Insert(selection.StartColumn - 1, OutputToken);
                    module.ReplaceLine(e.CodePane.Selection.StartLine, replacement);
                    e.CodePane.Selection = selection;
                    return replacement;
                }
            }
            return null;
        }
    }

    public class AutoCompleteClosingParenthese : AutoCompleteBase
    {
        public override string InputToken => "(";
        public override string OutputToken => ")";
    }
    public class AutoCompleteClosingString : AutoCompleteBase
    {
        public override string InputToken => "\"";
        public override string OutputToken => "\"";
    }
    public class AutoCompleteClosingBracket : AutoCompleteBase
    {
        public override string InputToken => "[";
        public override string OutputToken => "]";
    }
    public class AutoCompleteClosingBrace : AutoCompleteBase
    {
        public override string InputToken => "{";
        public override string OutputToken => "}";
    }
    public class AutoCompleteEndIf : AutoCompleteBase
    {
        public override string InputToken => "If ";
        public override string OutputToken => "End If";

        public override string Execute(TypingCodeEventArgs e)
        {
            var selection = e.CodePane.Selection;

            if (e.IsCommitted && e.Code.Trim().StartsWith(InputToken))
            {
                var indent = e.Code.IndexOf(InputToken + 1); // borked
                using (var module = e.CodePane.CodeModule)
                {
                    var code = OutputToken.PadLeft(indent + OutputToken.Length, ' ');
                    module.InsertLines(selection.StartLine + 1, code);
                    e.CodePane.Selection = selection; // todo auto-indent?
                    return code;
                }
            }
            return null;
        }
    }

    public interface ISelectionChangeService
    {
        event EventHandler<DeclarationChangedEventArgs> SelectedDeclarationChanged;
        event EventHandler<DeclarationChangedEventArgs> SelectionChanged;
    }
    
    public interface ITypingCodeService
    {
        event EventHandler TypingCode;
    }

    public class TypingCodeService : ITypingCodeService, IDisposable
    {
        public event EventHandler TypingCode;
        private readonly IReadOnlyList<IAutoComplete> _autocompletions = new IAutoComplete[]
        {
            new AutoCompleteClosingParenthese(),
            new AutoCompleteClosingString(),
            new AutoCompleteClosingBracket(),
            new AutoCompleteClosingBrace(),
            new AutoCompleteEndIf(),
        };

        public TypingCodeService()
        {
            VBENativeServices.TypingCode += VBENativeServices_TypingCode;
        }

        QualifiedSelection? _lastSelection;
        string _lastCode;

        private void VBENativeServices_TypingCode(object sender, TypingCodeEventArgs e)
        {
            TypingCode?.Invoke(this, e);
            var selection = e.CodePane.Selection;
            var qualifiedSelection = e.CodePane.GetQualifiedSelection();

            if (!selection.IsSingleCharacter || e.Code.Equals(_lastCode) || qualifiedSelection.Value.Equals(_lastSelection))
            {
                return;
            }

            foreach (var autocomplete in _autocompletions.Where(auto => auto.IsEnabled))
            {
                var replacement = autocomplete.Execute(e);
                if (replacement != null)
                {
                    _lastSelection = qualifiedSelection;
                    _lastCode = replacement;
                    break;
                }
            }
        }

        public void Dispose()
        {
            VBENativeServices.TypingCode -= VBENativeServices_TypingCode;
        }
    }

    public class SelectionChangeService : ISelectionChangeService, IDisposable
    {
        public event EventHandler<DeclarationChangedEventArgs> SelectedDeclarationChanged;
        public event EventHandler<DeclarationChangedEventArgs> SelectionChanged;

        private Declaration _lastSelectedDeclaration;
        private readonly IVBE _vbe;
        private readonly IParseCoordinator _parser;

        public SelectionChangeService(IVBE vbe, IParseCoordinator parser)
        {
            _parser = parser;
            _vbe = vbe;
            VBENativeServices.SelectionChanged += OnVbeSelectionChanged;
            VBENativeServices.WindowFocusChange += OnVbeFocusChanged;
        }
        
        private void OnVbeSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.CodePane == null || e.CodePane.IsWrappingNullReference)
            {
                return;
            }

            new Task(() =>
            {
                var eventArgs = new DeclarationChangedEventArgs(e.CodePane, _parser.State.FindSelectedDeclaration(e.CodePane));
                DispatchSelectedDeclaration(eventArgs);
            }).Start();
        }

        private void OnVbeFocusChanged(object sender, WindowChangedEventArgs e)
        {
            if (e.EventType == FocusType.GotFocus)
            {
                switch (e.Window.Type)
                {
                    case WindowKind.Designer:
                        //Designer or control on designer form selected.
                        if (e.Window == null || e.Window.IsWrappingNullReference || e.Window.Type != WindowKind.Designer)
                        {
                            return;
                        }
                        new Task(() => DispatchSelectedDesignerDeclaration(_vbe.SelectedVBComponent)).Start();                  
                        break;
                    case WindowKind.CodeWindow:
                        //Caret changed in a code pane.
                        if (e.CodePane != null && !e.CodePane.IsWrappingNullReference)
                        {
                            new Task(() => DispatchSelectedDeclaration(new DeclarationChangedEventArgs(e.CodePane, _parser.State.FindSelectedDeclaration(e.CodePane)))).Start(); 
                        }
                        break;
                }
            }
            else if (e.EventType == FocusType.ChildFocus)
            {
                //Treeview selection changed in project window.
                new Task(() => DispatchSelectedProjectNodeDeclaration(_vbe.SelectedVBComponent)).Start();
            }
        }

        private void DispatchSelectionChanged(DeclarationChangedEventArgs eventArgs)
        {
            SelectionChanged?.Invoke(null, eventArgs);
        }
       
        private void DispatchSelectedDeclaration(DeclarationChangedEventArgs eventArgs)
        {
            DispatchSelectionChanged(eventArgs);

            if (!DeclarationChanged(eventArgs.Declaration))
            {
                return;
            }

            _lastSelectedDeclaration = eventArgs.Declaration;

            SelectedDeclarationChanged?.Invoke(null, eventArgs);
        }

        private void DispatchSelectedDesignerDeclaration(IVBComponent component)
        {           
            if (string.IsNullOrEmpty(component?.Name))
            {
                return;
            }

            var selectedCount = component.SelectedControls.Count;
            if (selectedCount == 1)
            {
                var name = component.SelectedControls.Single().Name;
                var control =
                    _parser.State.DeclarationFinder.MatchName(name)
                        .SingleOrDefault(d => d.DeclarationType == DeclarationType.Control
                                              && d.ProjectId == component.ParentProject.ProjectId
                                              && d.ParentDeclaration.IdentifierName == component.Name);

                DispatchSelectedDeclaration(new DeclarationChangedEventArgs(null, control, component));            
                return;
            }
            var form =
                _parser.State.DeclarationFinder.MatchName(component.Name)
                    .SingleOrDefault(d => d.DeclarationType.HasFlag(DeclarationType.ClassModule)
                                          && d.ProjectId == component.ParentProject.ProjectId);

            DispatchSelectedDeclaration(new DeclarationChangedEventArgs(null, form, component, selectedCount > 1));
        }

        private void DispatchSelectedProjectNodeDeclaration(IVBComponent component)
        {
            if (_parser.State.DeclarationFinder == null)
            {
                return;
            }

            if ((component == null || component.IsWrappingNullReference) && !_vbe.ActiveVBProject.IsWrappingNullReference)
            {
                //The user might have selected the project node in Project Explorer. If they've chosen a folder, we'll return the project anyway.
                var project =
                    _parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                        .SingleOrDefault(decl => decl.ProjectId.Equals(_vbe.ActiveVBProject.ProjectId));

                DispatchSelectedDeclaration(new DeclarationChangedEventArgs(null, project, component));
            }
            else if (component != null && component.Type == ComponentType.UserForm && component.HasOpenDesigner)
            {
                DispatchSelectedDesignerDeclaration(component);
            }
            else if (component != null)
            {
                
                var module =
                    _parser.State.AllUserDeclarations.SingleOrDefault(
                        decl => decl.DeclarationType.HasFlag(DeclarationType.Module) &&
                                decl.IdentifierName.Equals(component.Name) &&
                                decl.ProjectId.Equals(_vbe.ActiveVBProject.ProjectId));

                DispatchSelectedDeclaration(new DeclarationChangedEventArgs(null, module, component));
            }
        }

        private bool DeclarationChanged(Declaration current)
        {
            if ((_lastSelectedDeclaration == null && current == null) ||
                ((_lastSelectedDeclaration != null && current != null) && _lastSelectedDeclaration.Equals(current)))
            {
                return false;
            }
            return true;
        }

        public void Dispose()
        {
            VBENativeServices.SelectionChanged -= OnVbeSelectionChanged;
            VBENativeServices.WindowFocusChange -= OnVbeFocusChanged;
        }
    }

    public class DeclarationChangedEventArgs : EventArgs
    {
        public ICodePane ActivePane { get; private set; }
        public Declaration Declaration { get; private set; }
        // ReSharper disable once InconsistentNaming
        public IVBComponent VBComponent { get; private set; }
        public bool MultipleControlsSelected { get; private set; }

        public DeclarationChangedEventArgs(ICodePane pane, Declaration declaration, IVBComponent component = null, bool multipleControls = false)
        {
            ActivePane = pane;
            Declaration = declaration;
            VBComponent = component;
            MultipleControlsSelected = multipleControls;
        }
    }
}
