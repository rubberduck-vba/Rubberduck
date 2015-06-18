using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.ToDoItems
{
    /// <summary>
    /// Presenter for the to-do items explorer.
    /// </summary>
    public class ToDoExplorerDockablePresenter : DockablePresenterBase
    {
        private readonly IRubberduckParser _parser;
        private readonly IEnumerable<ToDoMarker> _markers;
        private readonly GridViewSort<ToDoItem> _gridViewSort;
        private readonly IToDoExplorerWindow _view;

        public ToDoExplorerDockablePresenter(IRubberduckParser parser, IEnumerable<ToDoMarker> markers, VBE vbe, AddIn addin, IToDoExplorerWindow window, GridViewSort<ToDoItem> gridViewSort)
            : base(vbe, addin, window)
        {
            _parser = parser;
            _markers = markers;
            _gridViewSort = gridViewSort;

            _view = window;
            _view.NavigateToDoItem += NavigateToDoItem;
            _view.RefreshToDoItems += RefreshToDoList;
            _view.RemoveToDoMarker += RemoveMarker;
            _view.SortColumn += SortColumn;
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
                _view.TodoItems = await GetItems();
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

        private void RemoveMarker(object sender, EventArgs e)
        {
            var selectedIndex = _view.GridView.SelectedRows[0].Index;
            var dataSource = ((BindingList<ToDoItem>)_view.GridView.DataSource).ToList();
            var selectedItem = dataSource[selectedIndex];

            var module = selectedItem.GetSelection().QualifiedName.Component.CodeModule;

            var oldContent = module.Lines[selectedItem.LineNumber, 1];
            var newContent =
                oldContent.Remove(selectedItem.GetSelection().Selection.StartColumn - 1);

            module.ReplaceLine(selectedItem.LineNumber, newContent);

            Refresh();
        }

        private void SortColumn(object sender, DataGridViewCellMouseEventArgs e)
        {
            var columnName = _view.GridView.Columns[e.ColumnIndex].Name;

            _view.TodoItems = _gridViewSort.Sort(_view.TodoItems, columnName);
        }

        private async Task<IOrderedEnumerable<ToDoItem>> GetItems()
        {
            //await Task.Yield();

            var items = new ConcurrentBag<ToDoItem>();
            var projects = VBE.VBProjects.Cast<VBProject>().Where(project => project.Protection != vbext_ProjectProtection.vbext_pp_locked);
            foreach(var project in projects)
            {
                var modules = _parser.Parse(project, this).ComponentParseResults;
                foreach (var module in modules)
                {
                    var markers = module.Comments.SelectMany(GetToDoMarkers);
                    foreach (var marker in markers)
                    {
                        items.Add(marker);
                    }
                }
            };

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
                           .Select(marker => new ToDoItem(marker.Priority, comment));
        }

        private void NavigateToDoItem(object sender, ToDoItemClickEventArgs e)
        {
            var projects = VBE.VBProjects.Cast<VBProject>()
                .Where(p => p.Protection != vbext_ProjectProtection.vbext_pp_locked
                            && p.Name == e.SelectedItem.ProjectName
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
