using System;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AddRemoveReferences
{
    public class ReferenceModel
    {
        public ReferenceModel(IVBProject project, int priority)
        {
            Name = project.Name;
            Priority = priority;
            Guid = string.Empty;
            Description = project.Description;
            Version = default(Version);
            FullPath = project.FileName;
            IsBuiltIn = false;
            Type = ReferenceKind.Project;
        }

        public ReferenceModel(dynamic info) // todo figure out what type that should be 
        {
            Name = info.Name;
            Guid = info.Guid;
            Description = info.Description;
            Version = info.Version;
            FullPath = info.FullPath;
            IsBuiltIn = false;
            Type = ReferenceKind.TypeLibrary;
        }

        public ReferenceModel(IReference reference, int priority)
        {
            IsSelected = true;
            Priority = priority;
            Name = reference.Name;
            Guid = reference.Guid;
            Description = reference.Description;
            Version = new Version(reference.Major, reference.Minor);
            FullPath = reference.FullPath;
            IsBuiltIn = reference.IsBuiltIn;
            IsBroken = reference.IsBroken;
            Type = reference.Type;
        }

        public bool IsSelected { get; set; }
        public bool IsRemoved { get; set; }
        public int Priority { get; set; }

        public string Name { get; }
        public string Guid { get; }
        public string Description { get; }
        public Version Version { get; }
        public string FullPath { get; }
        public bool IsBuiltIn { get; }
        public bool IsBroken { get; }
        public ReferenceKind Type { get; }

        public ReferenceStatus Status => IsBuiltIn
            ? ReferenceStatus.BuiltIn
            : IsBroken
                ? ReferenceStatus.Broken
                : IsRemoved
                    ? ReferenceStatus.Removed
                    : IsSelected
                        ? ReferenceStatus.Loaded
                        : ReferenceStatus.None;
    }
}