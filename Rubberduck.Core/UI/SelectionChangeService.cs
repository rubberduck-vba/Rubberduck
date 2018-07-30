using System;
using System.Linq;
using System.Threading.Tasks;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.WindowsApi;

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
            new Task(() =>
            {
                using (var active = _vbe.ActiveCodePane)
                {
                    if (active == null)
                    {
                        return;
                    }
                    var eventArgs = new DeclarationChangedEventArgs(_parser.State.FindSelectedDeclaration(active));
                    DispatchSelectedDeclaration(eventArgs);
                }
            }).Start();
        }

        private void OnVbeFocusChanged(object sender, WindowChangedEventArgs e)
        {
            if (e.EventType == FocusType.GotFocus)
            {
                switch (e.Hwnd.ToWindowType())
                {
                    case WindowType.DesignerWindow:
                        new Task(() =>
                        {
                            using (var component = _vbe.SelectedVBComponent)
                            {
                                DispatchSelectedDesignerDeclaration(component);
                            }                           
                        }).Start();                  
                        break;
                    case WindowType.CodePane:
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
                new Task(() =>
                {
                    using (var component = _vbe.SelectedVBComponent)
                    {
                        DispatchSelectedProjectNodeDeclaration(component);
                    }
                }).Start();
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

            using (var selected = component.SelectedControls)
            using (var parent = component.ParentProject)
            {
                var selectedCount = selected.Count;
                if (selectedCount == 1)
                {
                    var name = selected.Single().Name;
                    var control =
                        _parser.State.DeclarationFinder.MatchName(name)
                            .SingleOrDefault(d => d.DeclarationType == DeclarationType.Control
                                                  && d.ProjectId == parent.ProjectId
                                                  && d.ParentDeclaration.IdentifierName == component.Name);

                    DispatchSelectedDeclaration(new DeclarationChangedEventArgs(control));
                    return;
                }
                var form =
                    _parser.State.DeclarationFinder.MatchName(component.Name)
                        .SingleOrDefault(d => d.DeclarationType.HasFlag(DeclarationType.ClassModule)
                                              && d.ProjectId == parent.ProjectId);

                DispatchSelectedDeclaration(new DeclarationChangedEventArgs(form, selectedCount > 1));
            }
        }

        private void DispatchSelectedProjectNodeDeclaration(IVBComponent component)
        {
            if (_parser.State.DeclarationFinder == null)
            {
                return;
            }

            using (var active = _vbe.ActiveVBProject)
            {
                if ((component == null || component.IsWrappingNullReference) && !active.IsWrappingNullReference)
                {
                    //The user might have selected the project node in Project Explorer. If they've chosen a folder, we'll return the project anyway.
                    var project =
                        _parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                            .SingleOrDefault(decl => decl.ProjectId.Equals(active.ProjectId));

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
                                    decl.ProjectId.Equals(active.ProjectId));

                    DispatchSelectedDeclaration(new DeclarationChangedEventArgs(module));
                }
            }
        }

        private bool DeclarationChanged(Declaration current)
        {
            return (_lastSelectedDeclaration != null || current != null) &&
                   (_lastSelectedDeclaration == null || current == null || !_lastSelectedDeclaration.Equals(current));
        }

        public void Dispose()
        {
            VBENativeServices.SelectionChanged -= OnVbeSelectionChanged;
            VBENativeServices.WindowFocusChange -= OnVbeFocusChanged;
        }
    }

    public class DeclarationChangedEventArgs : EventArgs
    {
        public Declaration Declaration { get; }
        public bool MultipleControlsSelected { get; }

        public DeclarationChangedEventArgs(Declaration declaration, bool multipleControls = false)
        {
            Declaration = declaration;
            MultipleControlsSelected = multipleControls;
        }
    }
}
