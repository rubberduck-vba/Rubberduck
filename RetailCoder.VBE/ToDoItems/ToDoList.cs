using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Rubberduck.ToDoItems
{
    [System.Runtime.InteropServices.ComVisible(false)]
    public class ToDoList : BindingList<ToDoItem>
    {
        private List<Config.ToDoMarker> markers;
        private VBE vbe;

        public ToDoList(VBE vbe, List<Config.ToDoMarker> markers)
        {
            this.vbe = vbe;
            this.markers = markers;
            Refresh();
        }

        public void Refresh()
        {
            this.ClearItems();

            foreach (VBComponent component in this.vbe.ActiveVBProject.VBComponents)
            {
                CodeModule module = component.CodeModule;
                for (var i = 1; i <= module.CountOfLines; i++)
                {
                    string line = module.get_Lines(i, 1);
                    Config.ToDoMarker marker;

                    if (TryGetMarker(line, out marker))
                    {
                        var priority = (TaskPriority)marker.priority;
                        this.Add(new ToDoItem(priority, line, module, i));
                    }
                }
            }
        }

        private bool TryGetMarker(string line, out Config.ToDoMarker result)
        {
            foreach (var marker in this.markers)
            {
                if (line.Contains(marker.text, StringComparison.OrdinalIgnoreCase))
                {
                    result = marker;
                    return true;
                }
            }
            result = null;
            return false;
        }
    }
}
