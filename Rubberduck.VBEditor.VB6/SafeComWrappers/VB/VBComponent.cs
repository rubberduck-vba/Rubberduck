using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBComponent : SafeComWrapper<VB.VBComponent>, IVBComponent
    {
        public VBComponent(VB.VBComponent target, bool rewrapping = false) 
            : base(target, rewrapping)
        { }

        public QualifiedModuleName QualifiedModuleName => new QualifiedModuleName(this);

        public ComponentType Type => IsWrappingNullReference ? 0 : (ComponentType)Target.Type;

        public bool HasCodeModule => Type != ComponentType.RelatedDocument && Type != ComponentType.ResFile;
        public ICodeModule CodeModule
        {
            get
            {                
                if (!IsWrappingNullReference && HasCodeModule)
                {
                    return new CodeModule(Target.CodeModule);
                }
                return new CodeModule(null);
            }
        }

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);
        public IVBComponents Collection => new VBComponents(IsWrappingNullReference ? null : Target.Collection);
        public IProperties Properties => new Properties(IsWrappingNullReference ? null : Target.Properties);
        public bool HasOpenDesigner => !IsWrappingNullReference && Target.HasOpenDesigner;
        public string DesignerId => IsWrappingNullReference ? string.Empty : Target.DesignerID;

        public string Name
        {
            get
            {
                if (IsWrappingNullReference)
                {
                    return string.Empty;
                }
                if (!string.IsNullOrEmpty(Target.Name))
                {
                    return Target.Name;
                }
                if (FileCount > 0)
                {
                    return GetFileName(1);
                }

                Debug.Assert(false, "Could not get component name");
                return string.Empty;
            }
            set
            {
                if (!IsWrappingNullReference)
                {
                    Target.Name = value;
                }
            }
        }

        private string SafeName => Path.GetInvalidFileNameChars().Aggregate(Name, (current, c) => current.Replace(c.ToString(), "_"));

        public IControls Controls
        {
            get
            {
                using (var designer = IsWrappingNullReference
                    ? null
                    : new UserForm(Target.Designer as VB.VBForm))
                {
                    return designer == null
                        ? new VBControls(null)
                        : designer.Controls;
                }
            }
        }

        public IControls SelectedControls
        {
            get
            {
                using (var designer = IsWrappingNullReference
                    ? null
                    : new UserForm(Target.Designer as VB.VBForm))
                {
                    return designer == null
                        ? new SelectedVBControls(null)
                        : designer.Selected;
                }
            }
        }
        
        public bool HasDesigner
        {
            get
            {
                if (IsWrappingNullReference)
                {
                    return false;
                }
                using (var designer = new UserForm(Target.Designer as VB.VBForm))
                {
                    return !designer.IsWrappingNullReference;
                }
            }
        }

        public IWindow DesignerWindow() => new Window(IsWrappingNullReference ? null : Target.DesignerWindow());
        public void Activate() => Target.Activate();
        public bool IsSaved => !IsWrappingNullReference && !Target.IsDirty;
        public void Export(string path) => Target.SaveAs(path);

        /// <summary>
        /// Exports the component to the folder. The file name matches the component name and file extension is based on the component's type.
        /// </summary>
        /// <param name="folder">Destination folder for the resulting source file.</param>
        /// <param name="isTempFile">True if a unique temp file name should be generated. WARNING: filenames generated with this flag are not persisted.</param>
        /// <param name="specialCaseDocumentModules">If reimport of a document file is required later, it has to receive special treatment.</param>
        public string ExportAsSourceFile(string folder, bool isTempFile = false, bool specialCaseDocumentModules = true)
        {
            throw new NotSupportedException("Export as source file is not supported in VB6");
        }

        public IVBProject ParentProject
        {
            get
            {
                using (var collection = Collection)
                {
                    return collection.Parent;
                }
            }
        }

        public int FileCount => IsWrappingNullReference ? 0 : Target.FileCount;

        public string GetFileName(short index)
        {
            if (IsWrappingNullReference)
            {
                return null;
            }
            if (index < 1 || index > FileCount) // 1-based indexing from VB
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return Target.FileNames[index];
        }

        public override bool Equals(ISafeComWrapper<VB.VBComponent> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IVBComponent other)
        {
            return Equals(other as SafeComWrapper<VB.VBComponent>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        public int ContentHash()
        {
            if (IsWrappingNullReference || !HasCodeModule && !HasDesigner)
            {
                return 0;
            }

            var hashes = new List<int>();

            using (var code = CodeModule)
            {
                hashes.Add(code?.ContentHash() ?? 0);
            }

            if (HasDesigner)
            {
                using (var controls = Controls)
                {
                    hashes.AddRange(controls.Select(control => control.Name.GetHashCode()));
                }
            }

            return HashCode.Compute(hashes);
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}