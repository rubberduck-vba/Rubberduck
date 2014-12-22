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
        private Configuration _config;
        private BindingList<CodeInspection> _inspections;

        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public CodeInspectionControl()
        {
            InitializeComponent();
        }

        //todo: only require a code inspection list
        public CodeInspectionControl(Configuration config)
            : this()
        {
            _config = config;
            _inspections = new BindingList<CodeInspection>(_config.UserSettings.CodeInspectinSettings.CodeInspections.ToList());
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = _inspections;

            var nameColumn = new DataGridViewTextBoxColumn();
            nameColumn.Name = "InspectionName";
            nameColumn.DataPropertyName = "Name";
            nameColumn.HeaderText = "Inspection Name";
            this.dataGridView1.Columns.Add(nameColumn);

            var typeColumn = new DataGridViewTextBoxColumn();
            typeColumn.Name = "InspectionType";
            typeColumn.DataPropertyName = "InspectionType";
            typeColumn.HeaderText = "Type";
            this.dataGridView1.Columns.Add(typeColumn);

            var severityColumn = new DataGridViewComboBoxColumn();
            severityColumn.Name = "InspectionSeverity";
            severityColumn.DataPropertyName = "Severity";
            severityColumn.HeaderText = "Severity";
            severityColumn.DataSource = Enum.GetValues(typeof(Inspections.CodeInspectionSeverity));
            this.dataGridView1.Columns.Add(severityColumn);

            //todo: sexify form
        }
    }
}
