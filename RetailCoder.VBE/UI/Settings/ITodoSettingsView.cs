using System;
using System.ComponentModel;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;

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
        /// Request to remove the marker at the SelectedIndex.
        /// </summary>
        event EventHandler RemoveMarker;

        /// <summary>
        /// Request to add the currently active marker to BindingList{TodoMarker}.
        /// </summary>
        event EventHandler AddMarker;

        /// <summary>
        /// Raised whenver ActiveMarkerPriority is changed.
        /// </summary>
        event EventHandler PriorityChanged;

    }
}
