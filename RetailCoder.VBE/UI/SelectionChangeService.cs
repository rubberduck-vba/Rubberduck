using System;
using System.Linq;
using System.Threading.Tasks;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

namespace Rubberduck.UI
{
    public interface ISelectionChangeService
    {
        event EventHandler<DeclarationChangedEventArgs> SelectedDeclarationChanged;
        event EventHandler<DeclarationChangedEventArgs> SelectionChanged;
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
            if (SelectionChanged == null)
            {
                return;
            }
            SelectionChanged.Invoke(null, eventArgs);
        }
       
        private void DispatchSelectedDeclaration(DeclarationChangedEventArgs eventArgs)
        {
            DispatchSelectionChanged(eventArgs);

            if (!DeclarationChanged(eventArgs.Declaration))
            {
                return;
            }

            _lastSelectedDeclaration = eventArgs.Declaration;
            if (SelectedDeclarationChanged != null)
            {
                SelectedDeclarationChanged.Invoke(null, eventArgs);             
            }
        }

        private void DispatchSelectedDesignerDeclaration(IVBComponent component)
        {           
            if (component == null || string.IsNullOrEmpty(component.Name))
            {
                return;
            }

            var selected = component.SelectedControls.Count;
            if (selected == 1)
            {
                var name = component.SelectedControls.First().Name;
                var control =
                    _parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Control).SingleOrDefault(decl =>
                            decl.IdentifierName.Equals(name) &&
                            decl.ParentDeclaration.IdentifierName.Equals(component.Name) &&
                            decl.ProjectId.Equals(component.ParentProject.ProjectId));

                DispatchSelectedDeclaration(new DeclarationChangedEventArgs(null, control, component));            
                return;
            }
            var form =
                _parser.State.DeclarationFinder.UserDeclarations(DeclarationType.UserForm).SingleOrDefault(decl =>
                    decl.IdentifierName.Equals(component.Name) &&
                    decl.ProjectId.Equals(component.ParentProject.ProjectId));

            DispatchSelectedDeclaration(new DeclarationChangedEventArgs(null, form, component, selected > 1));
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
                ((_lastSelectedDeclaration != null && current != null) && !_lastSelectedDeclaration.Equals(current)))
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
