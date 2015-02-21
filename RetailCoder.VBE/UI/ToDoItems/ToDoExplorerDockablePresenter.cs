using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.Extensions;
using Rubberduck.ToDoItems;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Presenter for the to-do items explorer.
    /// </summary>
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
            Control.SortColumn += SortColumn;

            RefreshToDoList(this, EventArgs.Empty);
        }

        void SortColumn(object sender, DataGridViewCellMouseEventArgs e)
        {
            var columnName = Control.GridView.Columns[e.ColumnIndex].Name;
            IOrderedEnumerable<ToDoItem> resortedItems = null;


            if (columnName == Control.SortedByColumn && Control.SortedAscending)
            {
                resortedItems = Control.TodoItems.OrderByDescending(x => x.GetType().GetProperty(columnName).GetValue(x));
                Control.SortedAscending = false;
            }
            else
            {
                resortedItems = Control.TodoItems.OrderBy(x => x.GetType().GetProperty(columnName).GetValue(x));
                Control.SortedByColumn = columnName;
                Control.SortedAscending = true;
            }

            Control.TodoItems = resortedItems;
        }

        private void RefreshToDoList(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                var getItems = new Task<IOrderedEnumerable<ToDoItem>>(() => GetItems());
                getItems.Start();

                Control.TodoItems = getItems.Result;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private IOrderedEnumerable<ToDoItem> GetItems()
        {
            var items = new ConcurrentBag<ToDoItem>();
            var projects = VBE.VBProjects.Cast<VBProject>();
            Parallel.ForEach(projects,
                project =>
                {
                    var modules = _parser.Parse(project);
                    foreach (var module in modules)
                    {
                        var markers = module.Comments.AsParallel().SelectMany(GetToDoMarkers);
                        foreach (var marker in markers)
                        {
                            items.Add(marker);
                        }
                    }
                });

            var sortedItems = items.OrderBy(item => item.ProjectName)
                                    .ThenBy(item => item.ModuleName)
                                    .ThenByDescending(item => item.Priority)
                                    .ThenBy(item => item.LineNumber);

            return sortedItems;
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
                .FirstOrDefault(p => p.Name == e.SelectedItem.ProjectName);

            if (project == null)
            {
                return;
            }

            var component = project.VBComponents.Cast<VBComponent>()
                .FirstOrDefault(c => c.Name == e.SelectedItem.ModuleName);

            if (component == null)
            {
                return;
            }

            var codePane = component.CodeModule.CodePane;

            codePane.SetSelection(e.SelectedItem.GetSelection().Selection);
            codePane.ForceFocus();
        }
    }
}
