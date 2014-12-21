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
            //this.dataGridView1.AutoGenerateColumns = true;
            this.dataGridView1.DataSource = _inspections;

            //todo: change severity to combo box


        }
    }
}
