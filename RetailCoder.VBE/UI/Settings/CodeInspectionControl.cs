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
    [System.Runtime.InteropServices.ComVisible(true)]
    public partial class CodeInspectionControl : UserControl
    {
        private BindingList<CodeInspection> _inspections;

        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public CodeInspectionControl()
        {
            InitializeComponent();
        }

        public CodeInspectionControl(List<CodeInspection> inspections)
            : this()
        {
            _inspections = new BindingList<CodeInspection>(inspections
                                                            .OrderBy(c => c.InspectionType.ToString())
                                                            .ThenBy(c => c.Name)
                                                            .ToList()
                                                            );
            
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = _inspections;

            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            var nameColumn = new DataGridViewTextBoxColumn();
            nameColumn.Name = "InspectionName";
            nameColumn.DataPropertyName = "Name";
            nameColumn.HeaderText = "Name";
            nameColumn.FillWeight = 150;
            nameColumn.ReadOnly = true;
            this.dataGridView1.Columns.Add(nameColumn);

            var typeColumn = new DataGridViewTextBoxColumn();
            typeColumn.Name = "InspectionType";
            typeColumn.DataPropertyName = "InspectionType";
            typeColumn.HeaderText = "Type";
            typeColumn.ReadOnly = true;
            this.dataGridView1.Columns.Add(typeColumn);

            var severityColumn = new DataGridViewComboBoxColumn();
            severityColumn.Name = "InspectionSeverity";
            severityColumn.DataPropertyName = "Severity";
            severityColumn.HeaderText = "Severity";
            severityColumn.DataSource = Enum.GetValues(typeof(Inspections.CodeInspectionSeverity));
            this.dataGridView1.Columns.Add(severityColumn);

        }
    }
}
