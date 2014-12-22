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

        public CodeInspectionControl(Configuration config)
            : this()
        {
            _config = config;
            _inspections = new BindingList<CodeInspection>(_config.UserSettings.CodeInspectinSettings.CodeInspections.ToList());
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = _inspections;

            var comboBoxCol = new DataGridViewComboBoxColumn();
            comboBoxCol.Name = "Severity";
            comboBoxCol.DataPropertyName = "Severity";
            comboBoxCol.HeaderText = "Severity";
            comboBoxCol.DataSource = Enum.GetValues(typeof(Inspections.CodeInspectionSeverity));

            this.dataGridView1.Columns.Add(comboBoxCol);

            //todo: change severity to combo box
        }
    }
}
