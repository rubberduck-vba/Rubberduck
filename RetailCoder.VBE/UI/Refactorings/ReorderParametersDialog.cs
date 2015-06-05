using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Refactorings.ReorderParameters;

namespace Rubberduck.UI.Refactorings
{
    public partial class ReorderParametersDialog : Form, IReorderParametersView
    {
        public List<Parameter> Parameters { get; set; }
        private Parameter _selectedItem;
        private Rectangle _dragBoxFromMouseDown;
        Point _startPoint;
        private int _newRowIndex;

        public ReorderParametersDialog()
        {
            InitializeComponent();
            InitializeCaptions();

            MethodParametersGrid.SelectionChanged += MethodParametersGrid_SelectionChanged;
            MethodParametersGrid.MouseMove += MethodParametersGrid_MouseMove;
            MethodParametersGrid.MouseDown += MethodParametersGrid_MouseDown;
            MethodParametersGrid.DragOver += MethodParametersGrid_DragOver;
            MethodParametersGrid.DragDrop += MethodParametersGrid_DragDrop;
        }

        private void InitializeCaptions()
        {
            OkButton.Text = RubberduckUI.OK;
            CancelButton.Text = RubberduckUI.CancelButtonText;
            Text = RubberduckUI.ReorderParamsDialog_Caption;
            TitleLabel.Text = RubberduckUI.ReorderParamsDialog_TitleText;
            InstructionsLabel.Text = RubberduckUI.ReorderParamsDialog_InstructionsLabelText;
            MoveUpButton.Text = RubberduckUI.ReorderParamsDialog_MoveUpButtonText;
            MoveDownButton.Text = RubberduckUI.ReorderParamsDialog_MoveDownButtonText;
        }

        private void MethodParametersGrid_SelectionChanged(object sender, EventArgs e)
        {
            SelectionChanged();
        }

        private void MethodParametersGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left &&
                _dragBoxFromMouseDown != Rectangle.Empty && !_dragBoxFromMouseDown.Contains(e.X, e.Y))
            {
                var dropEffect = MethodParametersGrid.DoDragDrop(
                        MethodParametersGrid.Rows[_newRowIndex],
                        DragDropEffects.Move);
            }
        }

        private void MethodParametersGrid_MouseDown(object sender, MouseEventArgs e)
        {
            _newRowIndex = MethodParametersGrid.HitTest(e.X, e.Y).RowIndex;

            if (_newRowIndex == -1)
            {
                _dragBoxFromMouseDown = Rectangle.Empty;
                return;
            }

            _startPoint = new Point(e.X, e.Y);

            AdjustDragRectangle(e, SystemInformation.DragSize);
        }

        private void MethodParametersGrid_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void MethodParametersGrid_DragDrop(object sender, DragEventArgs e)
        {
            var clientPoint = MethodParametersGrid.PointToClient(new Point(e.X, e.Y));

            if (e.Effect == DragDropEffects.Move && _newRowIndex != -1)
            {
                var rowIndexOfItemUnderMouse = MethodParametersGrid.HitTest(clientPoint.X, clientPoint.Y).RowIndex;

                if (rowIndexOfItemUnderMouse < 0)
                {
                    if (clientPoint.Y < _startPoint.Y)
                    {
                        rowIndexOfItemUnderMouse = 0;
                    }
                    else
                    {
                        rowIndexOfItemUnderMouse = Parameters.Count - 1;
                    }
                }

                var tmp = Parameters.ElementAt(_newRowIndex);
                Parameters.RemoveAt(_newRowIndex);
                Parameters.Insert(rowIndexOfItemUnderMouse, tmp);
                ReselectParameter();
            }
        }

        private void AdjustDragRectangle(MouseEventArgs eventArgs, Size dragSize)
        {
            _dragBoxFromMouseDown = new Rectangle(new Point(eventArgs.X - (dragSize.Width / 2), eventArgs.Y - (dragSize.Height / 2)), dragSize);
        }

        public void InitializeParameterGrid()
        {
            MethodParametersGrid.AutoGenerateColumns = false;
            MethodParametersGrid.Columns.Clear();
            MethodParametersGrid.DataSource = Parameters;
            MethodParametersGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
            MethodParametersGrid.MultiSelect = false;
            MethodParametersGrid.AllowUserToResizeRows = false;
            MethodParametersGrid.AllowDrop = true;
            MethodParametersGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            var column = new DataGridViewTextBoxColumn
            {
                Name = "Parameter",
                DataPropertyName = "FullDeclaration",
                HeaderText = RubberduckUI.Parameter,
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };

            MethodParametersGrid.Columns.Add(column);
            _selectedItem = Parameters[0];
        }

        private void OkButtonClick(object sender, EventArgs e)
        {
            OnOkButtonClicked();
        }

        public event EventHandler CancelButtonClicked;
        public void OnCancelButtonClicked()
        {
            Hide();
        }

        public event EventHandler OkButtonClicked;
        public void OnOkButtonClicked()
        {
            var handler = OkButtonClicked;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void MoveUpButtonClicked(object sender, EventArgs e)
        {
            if (MethodParametersGrid.SelectedRows.Count == 0)
            {
                return;
            }

            var selectedIndex = GetFirstSelectedRowIndex(0);
            SwapParameters(selectedIndex, selectedIndex - 1);

            ReselectParameter();
        }

        private void MoveDownButtonClicked(object sender, EventArgs e)
        {
            if (MethodParametersGrid.SelectedRows.Count == 0)
            {
                return;
            }

            var selectedIndex = GetFirstSelectedRowIndex(0);
            SwapParameters(selectedIndex, selectedIndex + 1);
            
            ReselectParameter();
        }

        private int GetFirstSelectedRowIndex(int index)
        {
            return MethodParametersGrid.SelectedRows[index].Index;
        }

        private void SwapParameters(int index1, int index2)
        {
            var tmp = Parameters[index1];
            Parameters[index1] = Parameters[index2];
            Parameters[index2] = tmp;
        }

        private void ReselectParameter()
        {
            MethodParametersGrid.Refresh();
            MethodParametersGrid.Rows
                                .Cast<DataGridViewRow>()
                                .Single(row => row.DataBoundItem == _selectedItem).Selected = true;

            SelectionChanged();
        }

        private void SelectionChanged()
        {
            _selectedItem = MethodParametersGrid.SelectedRows.Count == 0
                ? null
                : (Parameter)MethodParametersGrid.SelectedRows[0].DataBoundItem;

            MoveUpButton.Enabled = _selectedItem != null
                && MethodParametersGrid.SelectedRows[0].Index != 0;

            MoveDownButton.Enabled = _selectedItem != null
                && MethodParametersGrid.SelectedRows[0].Index != Parameters.Count - 1;
        }
    }
}
