using System;
using System.Collections.Generic;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public interface IAddTodoMarkerView
    {
        /// <summary>
        /// List of all current TodoMarkers.
        /// </summary>
        List<ToDoMarker> TodoMarkers { get; set; }

        /// <summary>
        /// Current text of new marker.
        /// </summary>
        string MarkerText { get; set; }

        /// <summary>
        /// Sets UI display based on validity of marker.
        /// </summary>
        bool IsValidMarker { get; set; }

        /// <summary>
        /// Request to add the currently active marker to BindingList{TodoMarker}.
        /// </summary>
        event EventHandler AddMarker;

        /// <summary>
        /// Cancel adding marker.
        /// </summary>
        event EventHandler Cancel;

        /// <summary>
        /// Cancel adding marker.
        /// </summary>
        event EventHandler TextChanged;

        void Show();
        void Hide();
        void Close();
    }
}
