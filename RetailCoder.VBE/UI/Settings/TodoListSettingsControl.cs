using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.Settings
{
    public partial class TodoListSettingsControl : UserControl
    {
        private TodoSettingView _view;

        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public TodoListSettingsControl()
        {
            InitializeComponent();
        }

        public TodoListSettingsControl(TodoSettingView view):this()
        {
            _view = view;
            this.tokenListBox.DataSource = _view.Markers;
        }

        private void TodoListSettingsControl_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
