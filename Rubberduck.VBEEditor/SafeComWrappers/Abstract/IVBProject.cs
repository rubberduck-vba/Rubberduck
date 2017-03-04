using System;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBProject : ISafeComWrapper, IEquatable<IVBProject>
    {
        IApplication Application { get; }
        IApplication Parent { get; }
        IVBE VBE { get; }
        IVBProjects Collection { get; }
        IReferences References { get; }
        IVBComponents VBComponents { get; }

        string ProjectId { get; }
        string Name { get; set; }
        string Description { get; set; }
        string HelpFile { get; set; }
        string FileName { get; }
        string BuildFileName { get; }
        bool IsSaved { get; }

        ProjectType Type { get; }
        EnvironmentMode Mode { get; }
        ProjectProtection Protection { get; }

        void AssignProjectId();
        void SaveAs(string fileName);
        void MakeCompiledFile();
        void ExportSourceFiles(string folder);
        string ProjectDisplayName { get; }

        IReadOnlyList<string> ComponentNames();
    }
}