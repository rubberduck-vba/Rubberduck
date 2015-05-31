using System;
using System.Windows.Forms;
using Rubberduck.Config;
using Rubberduck.Inspections;

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

            try
            {
                var rootNode = new TreeNode("Rubberduck") {ImageKey = "Ducky", SelectedImageKey = "Ducky"};
            
                var navNode = new TreeNode(RubberduckUI.Navigation) {ImageKey = "Navigation", SelectedImageKey = "Navigation"};
                var todoNode = navNode.Nodes.Add(RubberduckUI.TodoSettings_Caption);
                rootNode.Nodes.Add(navNode);

                var codeinspectionNode = new TreeNode(RubberduckUI.CodeInspections) { ImageKey = "CodeInspections", SelectedImageKey = "CodeInspections"};
                codeinspectionNode.Nodes.Add(new TreeNode(RubberduckUI.CodeInspectionSettings_CodeQualityIssues) { Name = CodeInspectionType.CodeQualityIssues.ToString() });
                codeinspectionNode.Nodes.Add(new TreeNode(RubberduckUI.CodeInspectionSettings_LanguageOpportunities) { Name = CodeInspectionType.LanguageOpportunities.ToString() });
                codeinspectionNode.Nodes.Add(new TreeNode(RubberduckUI.CodeInspectionSettings_MaintainabilityAndReadabilityIssues) { Name = CodeInspectionType.MaintainabilityAndReadabilityIssues.ToString() });
                rootNode.Nodes.Add(codeinspectionNode);

                settingsTreeView.Nodes.Add(rootNode);
                settingsTreeView.Nodes[0].ExpandAll();
            }
            catch (Exception exception)
            {
            }
        }

        public event EventHandler<TreeViewEventArgs> NodeSelected;
        private void settingsTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            var handler = NodeSelected;
            if (handler != null)
            {
                handler(sender, e);
            }
        }
    }
}
