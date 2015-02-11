using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.Extensions;
using Rubberduck.ToDoItems;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Presenter for the to-do items explorer.
    /// </summary>
    [ComVisible(false)]
    public class ToDoExplorerDockablePresenter : DockablePresenterBase
    {
        private readonly IRubberduckParser _parser;
        private readonly IEnumerable<ToDoMarker> _markers;
        private ToDoExplorerWindow Control { get { return UserControl as ToDoExplorerWindow; } }

        public ToDoExplorerDockablePresenter(IRubberduckParser parser, IEnumerable<ToDoMarker> markers, VBE vbe, AddIn addin) 
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
            RefreshAsync();
        }

        private async void RefreshAsync()
        {
            var items = new ConcurrentBag<ToDoItem>();
            var projects = VBE.VBProjects.Cast<VBProject>();
            Parallel.ForEach(projects,
                async project =>
                {
                    var modules = await _parser.ParseAsync(project);
                    foreach (var module in modules)
                    {
                        var markers = module.Comments.AsParallel().SelectMany(GetToDoMarkers);
                        foreach (var marker in markers)
                        {
                            items.Add(marker);
                        }
                    }
                });

            Control.SetItems(items);
        }

        private IEnumerable<ToDoItem> GetToDoMarkers(CommentNode comment)
        {
            return _markers.Where(marker => comment.Comment.ToLowerInvariant()
                                                   .Contains(marker.Text.ToLowerInvariant()))
                           .Select(marker => new ToDoItem((TaskPriority)marker.Priority, comment));
        }

        private void NavigateToDoItem(object sender, ToDoItemClickEventArgs e)
        {
            var project = VBE.VBProjects.Cast<VBProject>()
                .FirstOrDefault(p => p.Name == e.SelectedItem.Selection.QualifiedName.ProjectName);

            if (project == null)
            {
                return;
            }

            var component = project.VBComponents.Cast<VBComponent>()
                .FirstOrDefault(c => c.Name == e.SelectedItem.Selection.QualifiedName.ModuleName);

            if (component == null)
            {
                return;
            }

            var codePane = component.CodeModule.CodePane;

            codePane.SetSelection(e.SelectedItem.Selection.Selection);
            codePane.ForceFocus();
        }
    }
}
