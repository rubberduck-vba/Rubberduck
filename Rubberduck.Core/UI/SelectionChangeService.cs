using System;
using System.Linq;
using System.Threading.Tasks;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
        
        private void OnVbeSelectionChanged(object sender, EventArgs e)
        {
            //if (e.CodePane == null || e.CodePane.IsWrappingNullReference)
            //{
            //    return;
            //}

            new Task(() =>
            {
                //TODO
                var eventArgs = new DeclarationChangedEventArgs(_parser.State.FindSelectedDeclaration(_vbe.ActiveCodePane));
                DispatchSelectedDeclaration(eventArgs);
            }).Start();
        }

        private void OnVbeFocusChanged(object sender, WindowChangedEventArgs e)
        {
            if (e.EventType == FocusType.GotFocus)
            {
                switch (e.Hwnd.ToWindowType())
                {
                    case VBENativeServices.WindowType.DesignerWindow:
                        //Designer or control on designer form selected.
                        //if (e.Window == null || e.Window.IsWrappingNullReference || e.Window.Type != WindowKind.Designer)
                        //{
                        //    return;
                        //}
                        new Task(() => DispatchSelectedDesignerDeclaration(_vbe.SelectedVBComponent)).Start();                  
                        break;
                    case VBENativeServices.WindowType.CodePane:
                        //Caret changed in a code pane.
                        new Task(() =>
                        {
                            using (var pane = VBENativeServices.GetCodePaneFromHwnd(e.Hwnd))
                            {
                                DispatchSelectedDeclaration(
                                    new DeclarationChangedEventArgs(_parser.State.FindSelectedDeclaration(pane)));
                            }
                        }).Start();
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

                //DispatchSelectedDeclaration(new DeclarationChangedEventArgs(null, control, component));
                DispatchSelectedDeclaration(new DeclarationChangedEventArgs(control));
                return;
            }
            var form =
                _parser.State.DeclarationFinder.MatchName(component.Name)
                    .SingleOrDefault(d => d.DeclarationType.HasFlag(DeclarationType.ClassModule)
                                          && d.ProjectId == component.ParentProject.ProjectId);

            //DispatchSelectedDeclaration(new DeclarationChangedEventArgs(null, form, component, selectedCount > 1));
            DispatchSelectedDeclaration(new DeclarationChangedEventArgs(form, selectedCount > 1));
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

                DispatchSelectedDeclaration(new DeclarationChangedEventArgs(project));
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

                DispatchSelectedDeclaration(new DeclarationChangedEventArgs(module));
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
        //public ICodePane ActivePane { get; private set; }
        public Declaration Declaration { get; private set; }
        // ReSharper disable once InconsistentNaming
        //public IVBComponent VBComponent { get; private set; }
        public bool MultipleControlsSelected { get; private set; }

        //public DeclarationChangedEventArgs(ICodePane pane, Declaration declaration, IVBComponent component = null, bool multipleControls = false)
        //{
        //    ActivePane = pane;
        //    Declaration = declaration;
        //    VBComponent = component;
        //    MultipleControlsSelected = multipleControls;
        //}

        public DeclarationChangedEventArgs(Declaration declaration, bool multipleControls = false)
        {
            Declaration = declaration;
            MultipleControlsSelected = multipleControls;
        }
    }
}
