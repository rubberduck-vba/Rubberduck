using System;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VBA
{
    internal class ComIds : IComIds
    {
        private static readonly Guid _vBComponentsEventsGuid = new Guid("0002E116-0000-0000-C000-000000000046");
        private static readonly Guid _vBProjectsEventsGuid = new Guid("0002E190-0000-0000-C000-000000000046");
        private static readonly IComponentEventDispIds _componentEventDispIds = new ComponentEventDispIdsPrivate();
        private static readonly IProjectEventDispIds _projectEventDispIds = new ProjectEventDispIdsPrivate();

        public Guid VBComponentsEventsGuid => _vBComponentsEventsGuid;
        public Guid VBProjectsEventsGuid => _vBProjectsEventsGuid;
        public IComponentEventDispIds ComponentEventDispIds => _componentEventDispIds;
        public IProjectEventDispIds ProjectEventDispIds => _projectEventDispIds;

        private class ComponentEventDispIdsPrivate : IComponentEventDispIds
        {
            public int ItemAdded => 1;
            public int ItemRemoved => 2;
            public int ItemRenamed => 3;
            public int ItemSelected => 4;
            public int ItemActivated => 5;
            public int ItemReloaded => 6;
        }

        private class ProjectEventDispIdsPrivate : IProjectEventDispIds
        {
            public int ItemAdded => 1;
            public int ItemRemoved => 2;
            public int ItemRenamed => 3;
            public int ItemActivated => 4;
        }
    }
}
