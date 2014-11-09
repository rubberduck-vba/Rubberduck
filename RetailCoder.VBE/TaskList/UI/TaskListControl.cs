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
        private BindingList<Task> taskList;

        public TaskListControl(VBE vbe)
        {
            this.vbe = vbe;

            InitializeComponent();
            
            RefreshTaskList();
            InitializeGrid();

        }

        private void InitializeGrid()
        {
            taskListGridView.DataSource = taskList;
            taskListGridView.CellDoubleClick += RefreshGridView;

        }

        private void RefreshGridView(object sender, DataGridViewCellEventArgs e)
        {
            RefreshTaskList();
            taskListGridView.DataSource = taskList;
            taskListGridView.Refresh();
        }

        public void RefreshTaskList()
        {
            this.taskList = new BindingList<Task>();

            foreach (VBComponent component in this.vbe.ActiveVBProject.VBComponents)
            {
                CodeModule module = component.CodeModule;
                for (var i = 1; i <= module.CountOfLines; i++)
                {
                    string line = module.get_Lines(i, 1);
                    if (IsTaskComment(line))
                    {
                        var priority = GetTaskPriority(line);
                        this.taskList.Add(new Task(priority, line, module, i));
                    }
                }
            }
        }

        private TaskPriority GetTaskPriority(string line)
        {
            //todo: Create xml config file to allow user customization of tags
            var upCasedLine = line.ToUpper();
            if (upCasedLine.Contains("'BUG:"))
            {
                return TaskPriority.Bug;
            }

            return TaskPriority.Low;
        }


        private bool IsTaskComment(string line)
        {
            var upCasedLine = line.ToUpper();
            if (upCasedLine.Contains("'TODO:") || upCasedLine.Contains("'BUG:"))
            {
                return true;
            }

            return false;
        }
    }
}
