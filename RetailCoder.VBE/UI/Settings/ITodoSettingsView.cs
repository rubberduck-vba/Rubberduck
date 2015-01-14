using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Config;
using System.ComponentModel;
using System.Windows.Forms;

namespace Rubberduck.UI.Settings
{
    public interface ITodoSettingsView
    {
        TodoPriority ActiveMarkerPriority { get; set; }
        string ActiveMarkerText { get; set; }
        BindingList<ToDoMarker> TodoMarkers { get; set; }
        int SelectedIndex { get; set; }
        bool SaveEnabled { get; set; }

        event EventHandler RemoveMarker;
        event EventHandler AddMarker;
        event EventHandler SaveMarker;
        event EventHandler SelectionChanged;
        event EventHandler TextChanged;
        event EventHandler PriorityChanged;

    }
}
