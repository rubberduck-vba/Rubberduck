using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.Config;

namespace Rubberduck.UI.Settings
{
    public partial class CodeInspectionControl : UserControl
    {
        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public CodeInspectionControl()
        {
            InitializeComponent();
        }

        public CodeInspectionControl(IEnumerable<CodeInspection> inspections)
            : this()
        {
            var allInspections = new BindingList<CodeInspection>(inspections
                .OrderBy(c => c.InspectionType.ToString())
                .ThenBy(c => c.Name)
                .ToList()
                );
            
            codeInspectionsGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
            codeInspectionsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            codeInspectionsGrid.AutoGenerateColumns = false;
            codeInspectionsGrid.DataSource = allInspections;

            codeInspectionsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

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
            severityColumn.HeaderText = "Severity";
            severityColumn.DataSource = Enum.GetValues(typeof(Inspections.CodeInspectionSeverity));
            codeInspectionsGrid.Columns.Add(severityColumn);

        }
    }
}
