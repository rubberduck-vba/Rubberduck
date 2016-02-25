using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class SourceControlPanel : UserControl, ISourceControlView
    {
        public SourceControlPanel()
        {
            InitializeComponent();
        }

        public string ClassId
        {
            get { return "19A32FC9-4902-4385-9FE7-829D4F9C441D"; }
        }

        public string Caption
        {
            get { return RubberduckUI.SourceControlPanel_Caption; }
        }

        public ViewModelBase ViewModel
        {
            get { return (SourceControlViewViewModel)SourceControlPanelControl.DataContext; }
            set { SourceControlPanelControl.DataContext = value; }
        }
    }
}
