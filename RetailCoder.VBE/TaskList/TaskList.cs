using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RetailCoderVBE.Reflection;
using Microsoft.Vbe.Interop;

namespace RetailCoderVBE.TaskList
{
    class TaskList
    {

        private List<Task> _taskList;
        private VBE _vbe;

        public TaskList(VBE vbe)
        {
            _vbe = vbe;
            Refresh();
        }

        public void Refresh()
        {
            _taskList = new List<Task>();

            foreach (VBComponent component in _vbe.ActiveVBProject.VBComponents)
            {
                CodeModule module = component.CodeModule;
                for (var i = 1; i <= module.CountOfLines; i++ )
                {
                    string line = module.get_Lines(i,1);
                    if (IsTaskComment(line))
                    {
                        var priority = GetTaskPriority(line);
                        _taskList.Add(new Task(priority, line, module, i));
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
