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
    public partial class ConfigurationTreeView : UserControl
    {

        private Configuration _config;

        public ConfigurationTreeView(Configuration config)
        {
            InitializeComponent();

            _config = config;
        }

        private void InitializeTreeView()
        {
            var rootNode = new TreeNode("Rubberduck");
            var todoNode = rootNode.Nodes.Add("Todo List");
            var codeinspectionNode = rootNode.Nodes.Add("Code Inpsections");   
        }

        private void settingsTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }
    }
}
