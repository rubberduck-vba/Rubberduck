using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Rubberduck.Navigation.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// Interaction logic for CodeExplorerControl.xaml
    /// </summary>
    public partial class CodeExplorerControl
    {
        public CodeExplorerControl()
        {
            InitializeComponent();
        }

        private CodeExplorerViewModel ViewModel => DataContext as CodeExplorerViewModel;

        private void TreeView_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel != null && ViewModel.OpenCommand.CanExecute(ViewModel.SelectedItem))
            {
                ViewModel.OpenCommand.Execute(ViewModel.SelectedItem);
                e.Handled = true;
            }
        }

        private void TreeView_OnMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            ((TreeViewItem)sender).IsSelected = true;
            e.Handled = true;
        }

        private void ProjectTree_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                ViewModel.RemoveCommand.Execute(ViewModel.SelectedItem);
                e.Handled = true;
            }
        }

        private Point _lastLeftClickPosition;

        private void TreeView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _lastLeftClickPosition = e.GetPosition(null);
        }

        private void TreeView_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            EvaluateDragInitialization(sender, e);
        }

        private void EvaluateDragInitialization(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed
                || !(DataContext is CodeExplorerViewModel viewModel)
                || !viewModel.AllowDragAndDrop)
            {
                return;
            }
            
            var currentPosition = e.GetPosition(null);
            var fromLeftClick = currentPosition - _lastLeftClickPosition;

            if (Math.Abs(fromLeftClick.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(fromLeftClick.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            InitiateDrag(sender, e);
        }

        private void InitiateDrag(object sender, MouseEventArgs e)
        {
            if (ViewModel == null)
            {
                return;
            }

            EvaluateMoveToFolderDrag(sender, e);
        }

        private const string DragFolderMoveItemDataName = "dragFolderMoveItem";
        private void EvaluateMoveToFolderDrag(object sender, MouseEventArgs e)
        {
            var selectedItem = ViewModel.SelectedItem;
            if (!ViewModel.MoveToFolderDragAndDropCommand.CanExecute(("targetFolder", selectedItem)))
            {
                return;
            }

            var sourceTreeItem = (TreeViewItem)sender;
            var dragData = new DataObject(DragFolderMoveItemDataName, selectedItem);
            DragDrop.DoDragDrop(sourceTreeItem, dragData, DragDropEffects.Move);
        }

        private void TreeView_PreviewContinueDrag(object sender, QueryContinueDragEventArgs e)
        {
            if (e.EscapePressed)
            {
                e.Action = DragAction.Cancel;
            }
        }

        private void TreeView_OnDragOver(object sender, DragEventArgs e)
        {
            if (sender == e.OriginalSource)
            {
                return;
            }

            EvaluateCanDrop(sender, e);

            e.Handled = true;
        }

        private void EvaluateCanDrop(object sender, DragEventArgs e)
        {
            if (e.Handled)
            {
                return;
            }

            if (e.Data.GetDataPresent(DragFolderMoveItemDataName))
            {
                EvaluateCanDropFolderMove(sender, e);
            }
        }

        private void EvaluateCanDropFolderMove(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.None;

            if (e.Handled 
                || !e.Data.GetDataPresent(DragFolderMoveItemDataName) 
                || !TryGetFolderViewModel(sender, out var folderViewModel))
            {
                return;
            }

            var targetFolder = folderViewModel.FullPath;
            //We have to cast here to get the correct type parameter of the value type for the command.
            var draggedItem = e.Data.GetData(DragFolderMoveItemDataName) as ICodeExplorerNode;

            var moveToFolderDragCommand = ViewModel.MoveToFolderDragAndDropCommand;

            if (moveToFolderDragCommand.CanExecute((targetFolder, draggedItem)))
            {
                e.Effects = DragDropEffects.Move;
                e.Handled = true;
            }
        }

        private static bool TryGetFolderViewModel(object sender, out CodeExplorerCustomFolderViewModel folderViewModel)
        {
            if((sender as TreeViewItem)?.Header is CodeExplorerCustomFolderViewModel folder)
            {
                folderViewModel = folder;
                return true;
            }

            folderViewModel = null;
            return false;
        }

        private void TreeView_OnDrop(object sender, DragEventArgs e)
        {
            if (sender == e.OriginalSource)
            {
                return;
            }

            EvaluateMoveToFolderDrag(sender, e);

            e.Handled = true;
        }

        private void EvaluateMoveToFolderDrag(object sender, DragEventArgs e)
        {
            if (e.Handled
                || !e.Data.GetDataPresent(DragFolderMoveItemDataName)
                || !TryGetFolderViewModel(sender, out var folderViewModel))
            {
                return;
            }

            var targetFolder = folderViewModel.FullPath;
            //We have to cast here to get the correct type parameter of the value type for the command.
            var draggedItem = e.Data.GetData(DragFolderMoveItemDataName) as ICodeExplorerNode;

            var moveToFolderDragCommand = ViewModel.MoveToFolderDragAndDropCommand;

            if (moveToFolderDragCommand.CanExecute((targetFolder, draggedItem)))
            {
                moveToFolderDragCommand.Execute((targetFolder, draggedItem));
                e.Effects = DragDropEffects.Move;
                e.Handled = true;
            }
        }
    }
}
