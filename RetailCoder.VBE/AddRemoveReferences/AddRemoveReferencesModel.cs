using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AddRemoveReferences
{
    public class AddRemoveReferencesModel : ViewModelBase
    {
        public AddRemoveReferencesModel(IReadOnlyList<ReferenceModel> model)
        {
            AvailableReferences = model;
        }

        public IEnumerable<ReferenceModel> AvailableReferences { get; }
    }

    public class ReferenceModel
    {
        public ReferenceModel(IVBProject project)
        {
            Name = project.Name;
            Guid = string.Empty;
            Description = project.Description;
            Version = default(Version);
            FullPath = project.FileName;
            IsBuiltIn = false;
            Type = ReferenceKind.Project;
        }

        /// <summary>
        /// Creates a 
        /// </summary>
        /// <param name="info"></param>
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

        public ReferenceModel(IReference reference)
        {
            IsSelected = true;
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
        public bool IsPinned { get; set; }

        public string Name { get; }
        public string Guid { get; }
        public string Description { get; }
        public Version Version { get; }
        public string FullPath { get; }
        public bool IsBuiltIn { get; }
        public bool IsBroken { get; }
        public ReferenceKind Type { get; }
    }
}
