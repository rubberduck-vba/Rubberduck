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
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.AddRemoveReferences
{
    public class ReferenceModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public ReferenceModel(IVBProject project, int priority)
        {
            Name = project.Name ?? string.Empty;
            Priority = priority;
            Guid = Guid.Empty;
            Description = project.Description ?? string.Empty;
            FullPath = project.FileName ?? string.Empty;
            IsBuiltIn = false;
            Type = ReferenceKind.Project;
        }

        public ReferenceModel(RegisteredLibraryInfo info)
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

        public ReferenceModel(IReference reference, int priority)
        {
            Priority = priority;
            Name = reference.Name;
            Guid = new Guid(reference.Guid);
            Description = string.IsNullOrEmpty(reference.Description) ? Path.GetFileNameWithoutExtension(reference.FullPath) : reference.Description;
            Major = reference.Major;
            Minor = reference.Minor;
            FullPath = reference.FullPath;
            IsBuiltIn = reference.IsBuiltIn;
            IsBroken = reference.IsBroken;
            IsReferenced = true;
            Type = reference.Type;
        }

        public ReferenceModel(ITypeLib reference)
        {
            var documentation = new ComDocumentation(reference, -1);
            Name = documentation.Name;
            Description = documentation.DocString;

            reference.GetLibAttr(out var attributes);
            using (DisposalActionContainer.Create(attributes, reference.ReleaseTLibAttr))
            {
                var typeAttr = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.TYPELIBATTR>(attributes);

                Major = typeAttr.wMajorVerNum;
                Minor = typeAttr.wMinorVerNum;
                Flags = (TypeLibTypeFlags)typeAttr.wLibFlags;
                Guid = typeAttr.guid;
            }
        }

        public ReferenceModel(string path)
        {
            FullPath = path;
            Name = Path.GetFileName(path);
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

        public int? Priority { get; set; }
        
        public string Name { get; }
        public Guid Guid { get; }
        public string Description { get; }
        public string FullPath { get; }
        public string LocaleName { get; }

        public bool IsBuiltIn { get; }
        public bool IsBroken { get; }
        public TypeLibTypeFlags Flags { get; set;  }
        public ReferenceKind Type { get; }

        private string FullPath32 { get; }
        private string FullPath64 { get; }
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

        public ReferenceInfo ToReferenceInfo()
        {
            return new ReferenceInfo(Guid, Name, FullPath, Major, Minor);
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}