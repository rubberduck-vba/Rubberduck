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
        /// <summary>
        /// The TodoPriority level of the marker currently being edited.
        /// </summary>
        TodoPriority ActiveMarkerPriority { get; set; }

        /// <summary>
        /// The Text (or Name) of the marker currently being edited.
        /// </summary>
        string ActiveMarkerText { get; set; }

        /// <summary>
        /// List of all TodoMarkers to be displayed.
        /// </summary>
        BindingList<ToDoMarker> TodoMarkers { get; set; }

        /// <summary>
        /// Zero based index of the currently selected TodoMarker.
        /// </summary>
        int SelectedIndex { get; set; }

        /// <summary>
        /// Boolean value representing the enables/disabled state of the UI element the user needs to interact with to save the currently active marker.
        /// </summary>
        bool SaveEnabled { get; set; }

        /// <summary>
        /// Request to remove the marker at the SelectedIndex.
        /// </summary>
        event EventHandler RemoveMarker;

        /// <summary>
        /// Request to add the currently active marker to BindingList{TodoMarker}.
        /// </summary>
        event EventHandler AddMarker;

        /// <summary>
        /// Request to save changes made to the currently active marker back to the marker at the SelectedIndex.
        /// </summary>
        event EventHandler SaveMarker;

        /// <summary>
        /// Raised whenever SelectedIndex is changed.
        /// </summary>
        event EventHandler SelectionChanged;

        /// <summary>
        /// Raised whenever ActiveMarkerText is changed.
        /// </summary>
        event EventHandler TextChanged;

        /// <summary>
        /// Raised whenver ActiveMarkerPriority is changed.
        /// </summary>
        event EventHandler PriorityChanged;

    }
}
