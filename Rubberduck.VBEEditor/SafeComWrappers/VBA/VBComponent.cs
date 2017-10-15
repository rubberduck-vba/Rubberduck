using System;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBComponent : SafeComWrapper<VB.VBComponent>, IVBComponent
    {
        public VBComponent(VB.VBComponent target) 
            : base(target)
        {
        }

        public ComponentType Type
        {
            get { return IsWrappingNullReference ? 0 : (ComponentType)Target.Type; }
        }

        public ICodeModule CodeModule
        {
            get { return new CodeModule(IsWrappingNullReference ? null : Target.CodeModule); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public IVBComponents Collection
        {
            get { return new VBComponents(IsWrappingNullReference ? null : Target.Collection); }
        }

        public IProperties Properties
        {
            get { return new Properties(IsWrappingNullReference ? null : Target.Properties); }
        }

        public bool HasOpenDesigner
        {
            get { return !IsWrappingNullReference && Target.HasOpenDesigner; }
        }

        public string DesignerId
        {
            get { return IsWrappingNullReference ? string.Empty : Target.DesignerID; }
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Name; }
            set { if (!IsWrappingNullReference) Target.Name = value; }
        }

        private string SafeName
        {
            get { return Path.GetInvalidFileNameChars().Aggregate(Name, (current, c) => current.Replace(c.ToString(), "_")); }
        }

        public IControls Controls
        {
            get
            {
                var designer = IsWrappingNullReference
                    ? null
                    : Target.Designer as VB.Forms.UserForm;

                return designer == null 
                    ? new Controls(null) 
                    : new Controls(designer.Controls);
            }
        }

        public IControls SelectedControls
        {
            get
            {
                var designer = IsWrappingNullReference
                    ? null
                    : Target.Designer as VB.Forms.UserForm;

                return designer == null
                    ? new Controls(null)
                    : new Controls(designer.Selected);
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
                var designer = Target.Designer;
                var hasDesigner = designer != null;
                return hasDesigner;
            }
        }

        public IWindow DesignerWindow()
        {
            return new Window(IsWrappingNullReference ? null : Target.DesignerWindow());
        }

        public void Activate()
        {
            Target.Activate();
        }

        public bool IsSaved { get { return !IsWrappingNullReference && Target.Saved; } }

        public void Export(string path)
        {
            Target.Export(path);
        }

        /// <summary>
        /// Exports the component to the folder. The file is name matches the component name and file extension is based on the component's type.
        /// </summary>
        /// <param name="folder">Destination folder for the resulting source file.</param>
        /// <param name="tempFile">True if a unique temp file name should be generated. WARNING: filenames generated with this flag are not persisted.</param>
        public string ExportAsSourceFile(string folder, bool tempFile = false)
        {
            var fullPath = tempFile
                ? Path.Combine(folder, Path.GetRandomFileName())
                : Path.Combine(folder, SafeName + Type.FileExtension());
            switch (Type)
            {
                case ComponentType.UserForm:
                    ExportUserFormModule(fullPath);
                    break;
                case ComponentType.Document:
                    ExportDocumentModule(fullPath);
                    break;
                default:
                    Export(fullPath);
                    break;
            }

            return fullPath;
        }

        public IVBProject ParentProject
        {
            get { return Collection.Parent; }
        }

        private void ExportUserFormModule(string path)
        {
            // VBIDE API inserts an extra newline when exporting a UserForm module.
            // this issue causes forms to always be treated as "modified" in source control, which causes conflicts.
            // we need to remove the extra newline before the file gets written to its output location.

            var visibleCode = CodeModule.Content().Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var legitEmptyLineCount = visibleCode.TakeWhile(string.IsNullOrWhiteSpace).Count();

            var tempFile = ExportToTempFile();
            var contents = File.ReadAllLines(tempFile);
            var nonAttributeLines = contents.TakeWhile(line => !line.StartsWith("Attribute")).Count();
            var attributeLines = contents.Skip(nonAttributeLines).TakeWhile(line => line.StartsWith("Attribute")).Count();
            var declarationsStartLine = nonAttributeLines + attributeLines + 1;

            var emptyLineCount = contents.Skip(declarationsStartLine - 1)
                                         .TakeWhile(string.IsNullOrWhiteSpace)
                                         .Count();

            var code = contents;
            if (emptyLineCount > legitEmptyLineCount)
            {
                code = contents.Take(declarationsStartLine).Union(
                       contents.Skip(declarationsStartLine + emptyLineCount - legitEmptyLineCount))
                               .ToArray();
            }
            File.WriteAllLines(path, code);
        }

        private void ExportDocumentModule(string path)
        {
            var lineCount = CodeModule.CountOfLines;
            if (lineCount > 0)
            {
                var text = CodeModule.GetLines(1, lineCount);
                File.WriteAllText(path, text);
            }
        }

        private string ExportToTempFile()
        {
            var path = Path.Combine(Path.GetTempPath(), SafeName + Type.FileExtension());
            Export(path);
            return path;
        }
        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        DesignerWindow().Release();
        //        Controls.Release();
        //        Properties.Release();
        //        CodeModule.Release();
        //        base.Release(final);
        //    }
        //}

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
    }
}