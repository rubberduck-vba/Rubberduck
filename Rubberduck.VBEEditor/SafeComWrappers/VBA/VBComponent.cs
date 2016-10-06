using System;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBComponent : SafeComWrapper<Microsoft.Vbe.Interop.VBComponent>, IEquatable<VBComponent>
    {
        public VBComponent(Microsoft.Vbe.Interop.VBComponent comObject) 
            : base(comObject)
        {
        }

        public ComponentType Type
        {
            get { return IsWrappingNullReference ? 0 : (ComponentType)ComObject.Type; }
        }

        public CodeModule CodeModule
        {
            get { return new CodeModule(IsWrappingNullReference ? null : ComObject.CodeModule); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public VBComponents Collection
        {
            get { return new VBComponents(IsWrappingNullReference ? null : ComObject.Collection); }
        }

        public Properties Properties
        {
            get { return new Properties(IsWrappingNullReference ? null : ComObject.Properties); }
        }

        public bool HasOpenDesigner
        {
            get { return !IsWrappingNullReference && ComObject.HasOpenDesigner; }
        }

        public string DesignerId
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.DesignerID; }
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Name; }
            set { ComObject.Name = value; }
        }

        public Controls Controls
        {
            get
            {
                var designer = IsWrappingNullReference
                    ? null
                    : ComObject.Designer as Microsoft.Vbe.Interop.Forms.UserForm;

                return designer == null 
                    ? new Controls(null) 
                    : new Controls(designer.Controls);
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
                var designer = ComObject.Designer;
                var hasDesigner = designer != null;
                return hasDesigner;
            }
        }

        public Window DesignerWindow()
        {
            return new Window(IsWrappingNullReference ? null : ComObject.DesignerWindow());
        }

        public void Activate()
        {
            ComObject.Activate();
        }

        public bool IsSaved { get { return !IsWrappingNullReference && ComObject.Saved; } }

        public void Export(string path)
        {
            ComObject.Export(path);
        }

        /// <summary>
        /// Exports the component to the folder. The file is name matches the component name and file extension is based on the component's type.
        /// </summary>
        /// <param name="folder">Destination folder for the resulting source file.</param>
        public string ExportAsSourceFile(string folder)
        {
            var fullPath = Path.Combine(folder, Name + Type.FileExtension());
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
            var path = Path.Combine(Path.GetTempPath(), Name + Type.FileExtension());
            Export(path);
            return path;
        }
        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                DesignerWindow().Release();
                Controls.Release();
                Properties.Release();
                CodeModule.Release();
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.VBComponent> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(VBComponent other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBComponent>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}