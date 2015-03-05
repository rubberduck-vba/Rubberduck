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
        private IToDoExplorerWindow Control { get { return UserControl as IToDoExplorerWindow; } }

        public ToDoExplorerDockablePresenter(IRubberduckParser parser, IEnumerable<ToDoMarker> markers, VBE vbe, AddIn addin, IToDoExplorerWindow window) 
            : base(vbe, addin, window)
        {
            _parser = parser;
            _markers = markers;
            Control.NavigateToDoItem += NavigateToDoItem;
            Control.RefreshToDoItems += RefreshToDoList;
            Control.SortColumn += SortColumn;
        }

        public override void Show()
        {
            Refresh();
            base.Show();
        }

        public async void Refresh()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                Control.TodoItems = await GetItems();
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void RefreshToDoList(object sender, EventArgs e)
        {
            Refresh();
        }

        private void SortColumn(object sender, DataGridViewCellMouseEventArgs e)
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

        private async Task<IOrderedEnumerable<ToDoItem>> GetItems()
        {
            await Task.Yield();
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
            var projects = VBE.VBProjects.Cast<VBProject>()
                .Where(p => p.Name == e.SelectedItem.ProjectName
                                    && p.VBComponents.Cast<VBComponent>()
                                        .Any(c => c.Name == e.SelectedItem.ModuleName)
                                );

            if (projects == null)
            {
                return;
            }

            var component = projects.FirstOrDefault().VBComponents.Cast<VBComponent>()
                                    .First(c => c.Name == e.SelectedItem.ModuleName);

            if (component == null)
            {
                return;
            }

            var codePane = component.CodeModule.CodePane;
            codePane.SetSelection(e.SelectedItem.GetSelection().Selection);
        }
    }
}
