using System.Drawing;

namespace Rubberduck.UI.Commands
{
    public class ToolbarState
    {
        public static readonly Point UnsetLocation = new Point(-1, -1);
        public Point Location { get; set; }
        public bool Visible { get; set; }

        public ToolbarState() : this(UnsetLocation, false) { }

        public ToolbarState(Point location, bool isVisible)
        {
            Location = location;
            Visible = isVisible;
        }
    }
}