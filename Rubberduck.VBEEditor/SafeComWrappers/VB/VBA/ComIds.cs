using System;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VBA
{
    internal class VBComIds : IVBComIds
    {
        private static readonly Guid _VBComponentEvents = new Guid("0002E116-0000-0000-C000-000000000046");
        private static readonly Guid _VBProjectEvents = new Guid("0002E103-0000-0000-C000-000000000046");
        private static readonly IVBComponentEventDispIds _VBComponent = new VBComponentEventDispIds();
        private static readonly IVBProjectEventDispIds _VBProject = new VBProjectEventDispIds();

        public Guid VBComponentEvents => _VBComponentEvents;
        public Guid VBProjectEvents => _VBProjectEvents;
        public IVBComponentEventDispIds VBComponent => _VBComponent;
        public IVBProjectEventDispIds VBProject => _VBProject;

        private class VBComponentEventDispIds : IVBComponentEventDispIds
        {
            public int Added => 1;
            public int Removed => 2;
            public int Renamed => 3;
            public int Selected => 4;
            public int Activated => 5;
            public int Reloaded => 6;
        }

        private class VBProjectEventDispIds : IVBProjectEventDispIds
        {
            public int Added => 1;
            public int Removed => 2;
            public int Renamed => 3;
            public int Activated => 4;
        }
    }
}
