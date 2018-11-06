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
            VbeNativeServices.SelectionChanged += OnVbeSelectionChanged;
            VbeNativeServices.WindowFocusChange += OnVbeFocusChanged;
        }
        
        private void OnVbeSelectionChanged(object sender, EventArgs e)
        {
            Task.Run(() =>
            {
                using (var active = _vbe.ActiveCodePane)
                {
                    if (active == null)
                    {
                        return;
                    }
                    var eventArgs = new DeclarationChangedEventArgs(_vbe, _parser.State.FindSelectedDeclaration(active));
                    DispatchSelectedDeclaration(eventArgs);
                }
            });
        }

        private void OnVbeFocusChanged(object sender, WindowChangedEventArgs e)
        {
            if (e.EventType == FocusType.GotFocus)
            {
                switch (e.Hwnd.ToWindowType())
                {
                    case WindowType.DesignerWindow:
                        Task.Run(() =>
                        {
                            using (var component = _vbe.SelectedVBComponent)
                            {
                                DispatchSelectedDesignerDeclaration(component);
                            }                           
                        });                  
                        break;
                    case WindowType.CodePane:
                        //Caret changed in a code pane.
                        Task.Run(() =>
                        {
                            using (var pane = VbeNativeServices.GetCodePaneFromHwnd(e.Hwnd))
                            {
                                DispatchSelectedDeclaration(
                                    new DeclarationChangedEventArgs(_vbe, _parser.State.FindSelectedDeclaration(pane)));
                            }
                        });
                        break;
                }
            }
            else if (e.EventType == FocusType.ChildFocus)
            {
                //Treeview selection changed in project window.
                Task.Run(() =>
                {
                    using (var component = _vbe.SelectedVBComponent)
                    {
                        DispatchSelectedProjectNodeDeclaration(component);
                    }
                });
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

                    DispatchSelectedDeclaration(new DeclarationChangedEventArgs(_vbe, control));
                    return;
                }
                var form =
                    _parser.State.DeclarationFinder.MatchName(component.Name)
                        .SingleOrDefault(d => d.DeclarationType.HasFlag(DeclarationType.ClassModule)
                                              && d.ProjectId == parent.ProjectId);

                DispatchSelectedDeclaration(new DeclarationChangedEventArgs(_vbe, form, selectedCount > 1));
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

                    DispatchSelectedDeclaration(new DeclarationChangedEventArgs(_vbe, project));
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

                    DispatchSelectedDeclaration(new DeclarationChangedEventArgs(_vbe, module));
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
            VbeNativeServices.SelectionChanged -= OnVbeSelectionChanged;
            VbeNativeServices.WindowFocusChange -= OnVbeFocusChanged;
        }
    }

    public class DeclarationChangedEventArgs : EventArgs
    {
        public Declaration Declaration { get; }
        public string FallbackCaption { get; }
        public bool MultipleControlsSelected { get; }

        public DeclarationChangedEventArgs(IVBE vbe, Declaration declaration, bool multipleControls = false)
        {
            Declaration = declaration;
            MultipleControlsSelected = multipleControls;
            if (Declaration != null && !string.IsNullOrEmpty(Declaration.QualifiedName.MemberName))
            {
                return;
            }

            using (var active = vbe.SelectedVBComponent)
            using (var parent = active?.ParentProject)
            {
                FallbackCaption =
                    $"{parent?.Name ?? string.Empty}.{active?.Name ?? string.Empty} ({active?.Type.ToString() ?? string.Empty})"
                        .Trim('.');
            }
        }
    }
}
