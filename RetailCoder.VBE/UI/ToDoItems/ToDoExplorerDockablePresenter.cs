using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.Extensions;
using Rubberduck.ToDoItems;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>   (Not COM visible) Presenter for the Todo Explorer.  </summary>
    [ComVisible(false)]
    public class ToDoExplorerDockablePresenter : DockablePresenterBase
    {
        private readonly Parser _parser;
        private readonly IEnumerable<ToDoMarker> _markers;
        private ToDoExplorerWindow Control { get { return UserControl as ToDoExplorerWindow; } }

        public ToDoExplorerDockablePresenter(Parser parser, IEnumerable<ToDoMarker> markers, VBE vbe, AddIn addin) 
            : base(vbe, addin, new ToDoExplorerWindow())
        {
            _parser = parser;
            _markers = markers;
            Control.NavigateToDoItem += NavigateToDoItem;
            Control.RefreshToDoItems += RefreshToDoList;

            RefreshToDoList(this, EventArgs.Empty);
        }

        private void RefreshToDoList(object sender, EventArgs e)
        {
            var items = new List<ToDoItem>();
            foreach (var project in VBE.VBProjects.Cast<VBProject>())
            {
                var tree = _parser.Parse(project);
                items.AddRange(tree.FindAllComments().SelectMany(GetToDoMarkers));
            }

            Control.SetItems(items);
        }

        private IEnumerable<ToDoItem> GetToDoMarkers(Instruction instruction)
        {
            return _markers.Where(marker => instruction.Comment.ToLowerInvariant()
                                                   .Contains(marker.Text.ToLowerInvariant()))
                           .Select(marker => new ToDoItem((TaskPriority)marker.Priority, instruction));
        }

        private void NavigateToDoItem(object sender, ToDoItemClickEventArgs e)
        {
            var project = VBE.VBProjects.Cast<VBProject>()
                .FirstOrDefault(p => p.Name == e.Selection.ProjectName);

            if (project == null)
            {
                return;
            }

            var component = project.VBComponents.Cast<VBComponent>()
                .FirstOrDefault(c => c.Name == e.Selection.ModuleName);

            if (component == null)
            {
                return;
            }

            var codePane = component.CodeModule.CodePane;

            codePane.SetSelection(e.Selection.LineNumber);
            codePane.ForceFocus();
        }
    }
}
