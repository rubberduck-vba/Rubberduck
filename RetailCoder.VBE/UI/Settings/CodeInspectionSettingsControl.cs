using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Inspections;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public partial class CodeInspectionSettingsControl : UserControl
    {
        private GridViewSort<CodeInspectionSetting> _gridViewSort;

        private BindingList<CodeInspectionSetting> _allInspections;
        public BindingList<CodeInspectionSetting> AllInspections
        {
            get { return _allInspections; }
            set
            {
                _allInspections = value;
                FillGrid();
            }
        }

        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public CodeInspectionSettingsControl()
        {
            InitializeComponent();
        }

        public CodeInspectionSettingsControl(IEnumerable<CodeInspectionSetting> inspections, GridViewSort<CodeInspectionSetting> gridViewSort)
            : this()
        {
            _gridViewSort = gridViewSort;

            codeInspectionsGrid.AutoGenerateColumns = false;
            
            codeInspectionsGrid.BorderStyle = BorderStyle.None;
            codeInspectionsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            codeInspectionsGrid.CellBorderStyle = DataGridViewCellBorderStyle.None;
            codeInspectionsGrid.GridColor = Color.LightGray;

            codeInspectionsGrid.RowsDefaultCellStyle.BackColor = Color.White;
            codeInspectionsGrid.RowsDefaultCellStyle.SelectionBackColor = Color.LightBlue;
            codeInspectionsGrid.RowsDefaultCellStyle.SelectionForeColor = Color.MediumBlue;

            codeInspectionsGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.Honeydew;
            codeInspectionsGrid.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.LightBlue;
            codeInspectionsGrid.AlternatingRowsDefaultCellStyle.SelectionForeColor = Color.MediumBlue;

            AllInspections = new BindingList<CodeInspectionSetting>(inspections
                                 .OrderBy(c => c.Description)
                                 .ToList());

            // temporal coupling here: this code should run after columns are formatted.
            //codeInspectionsGrid.DataSource = AllInspections;
            codeInspectionsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            codeInspectionsGrid.ColumnHeaderMouseClick += SortColumn;
        }

        private void FillGrid()
        {
            codeInspectionsGrid.Columns.Clear();

            var nameColumn = new DataGridViewTextBoxColumn();
            nameColumn.Name = "Name";
            nameColumn.DataPropertyName = "Description";
            nameColumn.HeaderText = RubberduckUI.Name;
            nameColumn.FillWeight = 150;
            nameColumn.ReadOnly = true;
            codeInspectionsGrid.Columns.Add(nameColumn);

            var severityColumn = new DataGridViewComboBoxColumn();
            severityColumn.Name = "SeverityLabel";
            severityColumn.DataPropertyName = "SeverityLabel";
            severityColumn.DataSource = SettingsLabels();
            severityColumn.HeaderText = RubberduckUI.Severity;
            severityColumn.DefaultCellStyle.Font = codeInspectionsGrid.Font;
            codeInspectionsGrid.Columns.Add(severityColumn);

            codeInspectionsGrid.DataSource = AllInspections;
        }

        private List<string> SettingsLabels()
        {
            return (from object severity in Enum.GetValues(typeof (CodeInspectionSeverity))
                    select
                    RubberduckUI.ResourceManager.GetString("CodeInspectionSeverity_" + severity, RubberduckUI.Culture))
                    .ToList();
        }

        private void SortColumn(object sender, DataGridViewCellMouseEventArgs e)
        {
            var columnName = codeInspectionsGrid.Columns[e.ColumnIndex].Name;
            AllInspections = new BindingList<CodeInspectionSetting>(_gridViewSort.Sort(AllInspections.AsEnumerable(), columnName).ToList());
        }
    }
}
