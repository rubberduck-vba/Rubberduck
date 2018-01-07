using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Settings
{
    public class DefaultToDoMarkers
    {
        public IEnumerable<ToDoMarker> Markers { get; }

        public DefaultToDoMarkers()
        {
            var markersProperties = typeof(Properties.Settings).GetProperties().Where(prop => prop.PropertyType == typeof(ToDoMarker));

            Markers = markersProperties.Select(prop => prop.GetValue(Properties.Settings.Default)).Cast<ToDoMarker>();
        }
    }
}
