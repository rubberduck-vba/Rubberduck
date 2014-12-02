using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Config;

namespace Rubberduck.UI.Settings
{
    public class TodoSettingView
    {
        private List<IToDoMarker> _markers;
        public List<IToDoMarker> Markers { get { return _markers; } }

        public TodoSettingView(List<IToDoMarker> markers)
        {
            _markers = markers;
        }
    }
}
