using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace RetailCoderVBE.TaskList
{
    public partial class TaskListControl : UserControl
    {
        private VBE vbe;

        public TaskListControl(VBE vbe)
        {
            //todo: implement an actual task list control instead of this example
            this.BackColor = Color.Red;
            this.vbe = vbe;

            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("ToolWindow shown in VBA editor version " + vbe.Version);
        }

       
    }
}
