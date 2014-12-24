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
    public partial class ConfigurationTreeViewControl : UserControl
    {

        private Configuration _config;

        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public ConfigurationTreeViewControl()
        {
            InitializeComponent();
        }

        public ConfigurationTreeViewControl(Configuration config) : this()
        {
            _config = config;
            InitializeTreeView();
        }

        private void InitializeTreeView()
        {
            var rootNode = new TreeNode("Rubberduck");
            var todoNode = rootNode.Nodes.Add("Todo List");
            var codeinspectionNode = rootNode.Nodes.Add("Code Inpsections");   

            this.settingsTreeView.Nodes.Add(rootNode);
            this.settingsTreeView.Nodes[0].ExpandAll();
            
        }

        public event EventHandler<TreeViewEventArgs> NodeSelected;
        private void settingsTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //re-raise event
            NodeSelected(sender, e);
        }
    }
}
