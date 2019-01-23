using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AddRemoveReferences
{
    public class ReferenceModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        
        private ReferenceModel()
        {
            _info = new Lazy<ReferenceInfo>(GenerateInfo);
        }

        public ReferenceModel(ReferenceInfo info, ReferenceKind type, bool recent = false, bool pinned = false) : this()
        {
            Guid = info.Guid;
            Name = info.Name;
            Description = Name;
            FullPath = info.FullPath;
            Major = info.Major;
            Minor = info.Minor;
            IsRecent = recent;
            IsPinned = pinned;
            Type = type;
        }

        public ReferenceModel(IVBProject project, int priority) : this()
        {
            Name = project.Name ?? string.Empty;
            Priority = priority;
            Guid = Guid.Empty;
            Description = project.Description ?? project.Name;
            FullPath = project.FileName ?? string.Empty;
            IsBuiltIn = false;
            Type = ReferenceKind.Project;            
        }

        public ReferenceModel(RegisteredLibraryInfo info) : this()
        {
            Name = info.Name ?? string.Empty;
            Guid = info.Guid;
            Description = string.IsNullOrEmpty(info.Description) ? Path.GetFileNameWithoutExtension(info.FullPath) : info.Description;
            Major = info.Major;
            Minor = info.Minor;
            FullPath = info.FullPath;
            LocaleName = info.LocaleName;
            IsBuiltIn = false;
            Type = ReferenceKind.TypeLibrary;
            Flags = (TypeLibTypeFlags)info.Flags;
            IsRegistered = true;
        }

        public ReferenceModel(RegisteredLibraryInfo info, IReference reference, int priority) : this(info)
        {
            Priority = priority;
            IsBuiltIn = reference.IsBuiltIn;
            IsBroken = reference.IsBroken;
            IsReferenced = true;
        }

        public ReferenceModel(IReference reference, int priority) : this()
        {
            Priority = priority;
            Name = reference.Name;
            Guid = Guid.TryParse(reference.Guid, out var guid) ? guid : Guid.Empty;
            Description = string.IsNullOrEmpty(reference.Description) ? Path.GetFileNameWithoutExtension(reference.FullPath) : reference.Description;
            Major = reference.Major;
            Minor = reference.Minor;
            FullPath = reference.FullPath;
            IsBuiltIn = reference.IsBuiltIn;
            IsBroken = reference.IsBroken;
            IsReferenced = true;
            Type = reference.Type;
        }

        public ReferenceModel(string path, ITypeLib reference, IComLibraryProvider provider) : this()
        {
            FullPath = path;

            var documentation = provider.GetComDocumentation(reference);
            Name = documentation.Name;
            Description = documentation.DocString;

            var info = provider.GetReferenceInfo(reference, Name, path);
            Guid = info.Guid;
            Major = info.Major;
            Minor = info.Minor; 
        }

        public ReferenceModel(string path, bool broken = false) : this()
        {
            FullPath = path;
            try
            {
                Name = Path.GetFileName(path) ?? path;
                Description = Name;
            }
            catch
            {
                // Yeah, that's probably busted.
                IsBroken = true;
                return;
            }
            
            IsBroken = broken;
        }

        private bool _pinned;
        public bool IsPinned
        {
            get => _pinned;
            set
            {
                _pinned = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsRecent { get; set; }
        public bool IsRegistered { get; set; }
        public bool IsReferenced { get; set; }
        public bool IsUsed { get; set; }

        public int? Priority { get; set; }
        
        public string Name { get; } = string.Empty;
        public Guid Guid { get; }
        public string Description { get; } = string.Empty;
        public string FullPath { get; }
        public string LocaleName { get; } = string.Empty;

        public bool IsBuiltIn { get; set; }
        public bool IsBroken { get; }
        public TypeLibTypeFlags Flags { get; set;  }
        public ReferenceKind Type { get; }

        private string FullPath32 { get; } = string.Empty;
        private string FullPath64 { get; } = string.Empty;
        public int Major { get; set; }
        public int Minor { get; set; }
        public string Version => $"{Major}.{Minor}";

        public ReferenceStatus Status
        {
            get
            {
                var status = IsPinned ? ReferenceStatus.Pinned : ReferenceStatus.None;
                if (!Priority.HasValue)
                {
                    return IsRecent ? status | ReferenceStatus.Recent : status;
                }

                if (IsBroken)
                {
                    return status | ReferenceStatus.Broken;
                }

                if (IsBuiltIn)
                {
                    return status | ReferenceStatus.BuiltIn;
                }

                return status | (IsReferenced ? ReferenceStatus.Loaded : ReferenceStatus.Added);
            }
        }

        private readonly Lazy<ReferenceInfo> _info;
        private ReferenceInfo GenerateInfo() => new ReferenceInfo(Guid, Name, FullPath, Major, Minor);
        public ReferenceInfo ToReferenceInfo() => _info.Value;

        public bool Matches(ReferenceInfo info)
        {
            return Major == info.Major && Minor == info.Minor &&
                   FullPath.Equals(info.FullPath, StringComparison.OrdinalIgnoreCase) ||
                   FullPath32.Equals(info.FullPath, StringComparison.OrdinalIgnoreCase) ||
                   FullPath64.Equals(info.FullPath, StringComparison.OrdinalIgnoreCase) ||
                   !Guid.Equals(Guid.Empty) && Guid.Equals(info.Guid);
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}