using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Config;
using Rubberduck.Inspections;
using Rubberduck.UI.CodeInspections;

namespace Rubberduck.UI.Settings
{
    public partial class CodeInspectionSettingsControl : UserControl
    {
        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public CodeInspectionSettingsControl()
        {
            InitializeComponent();
        }

        public CodeInspectionSettingsControl(IEnumerable<CodeInspectionSetting> inspections)
            : this()
        {
            var allInspections = new BindingList<CodeInspectionSetting>(inspections
                .OrderBy(c => c.InspectionType.ToString())
                .ThenBy(c => c.Name)
                .ToList()
                );

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

            var nameColumn = new DataGridViewTextBoxColumn();
            nameColumn.Name = "InspectionName";
            nameColumn.DataPropertyName = "Name";
            nameColumn.HeaderText = "Name";
            nameColumn.FillWeight = 150;
            nameColumn.ReadOnly = true;
            codeInspectionsGrid.Columns.Add(nameColumn);

            var typeColumn = new DataGridViewTextBoxColumn();
            typeColumn.Name = "InspectionType";
            typeColumn.DataPropertyName = "InspectionType";
            typeColumn.HeaderText = "Type";
            typeColumn.ReadOnly = true;
            codeInspectionsGrid.Columns.Add(typeColumn);

            var severityColumn = new DataGridViewComboBoxColumn();
            severityColumn.Name = "InspectionSeverity";
            severityColumn.DataPropertyName = "Severity";
            severityColumn.DataSource = Enum.GetValues(typeof(CodeInspectionSeverity));
            severityColumn.HeaderText = "Severity";
            severityColumn.DefaultCellStyle.Font = codeInspectionsGrid.Font;
            codeInspectionsGrid.Columns.Add(severityColumn);

            // temporal coupling here: this code should run after columns are formatted.
            codeInspectionsGrid.DataSource = allInspections;
            codeInspectionsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
    }
}
