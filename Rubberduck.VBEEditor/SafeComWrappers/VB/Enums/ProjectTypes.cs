using System.Collections.Generic;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public static class ProjectTypes
    {
        public static ProjectType[] All { get; } = { ProjectType.StandardExe, ProjectType.ActiveXExe, ProjectType.ActiveXDll, ProjectType.ActiveXControl, ProjectType.HostProject, ProjectType.StandAlone };

        public static ProjectType[] VB6 { get; } = { ProjectType.StandardExe, ProjectType.ActiveXExe, ProjectType.ActiveXDll, ProjectType.ActiveXControl };

        public static ProjectType[] VBA { get; } = { ProjectType.HostProject, ProjectType.StandAlone };
    }
}
