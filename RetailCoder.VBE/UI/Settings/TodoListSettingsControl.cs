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

        public TodoListSettingsControl(TodoSettingView view)
        {
            InitializeComponent();

            _view = view;
            this.tokenListBox.DataSource = _view.Markers;
        }

        private void TodoListSettingsControl_Load(object sender, EventArgs e)
        {

        }
    }
}
