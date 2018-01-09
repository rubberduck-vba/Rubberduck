using System;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBComponent : SafeComWrapper<VB.VBComponent>, IVBComponent
    {
        public VBComponent(VB.VBComponent target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public ComponentType Type => IsWrappingNullReference ? 0 : (ComponentType)Target.Type;

        public ICodeModule CodeModule => new CodeModule(IsWrappingNullReference ? null : Target.CodeModule);

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IVBComponents Collection => new VBComponents(IsWrappingNullReference ? null : Target.Collection);

        public IProperties Properties => new Properties(IsWrappingNullReference ? null : Target.Properties);

        public bool HasOpenDesigner => !IsWrappingNullReference && Target.HasOpenDesigner;

        public string DesignerId => IsWrappingNullReference ? string.Empty : Target.DesignerID;

        public string Name
        {
            get => IsWrappingNullReference ? string.Empty : Target.Name;
            set => Target.Name = value;
        }

        public IControls Controls => throw new NotImplementedException();

        public IControls SelectedControls => throw new NotImplementedException();

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

        public bool IsSaved => throw new NotImplementedException();

        public void Export(string path)
        {
            //Target.Export(path);
        }

        /// <summary>
        /// Exports the component to the folder. The file name matches the component name and file extension is based on the component's type.
        /// </summary>
        /// <param name="folder">Destination folder for the resulting source file.</param>
        /// <param name="tempFile">True if a unique temp file name should be generated. WARNING: filenames generated with this flag are not persisted.</param>
        public string ExportAsSourceFile(string folder, bool tempFile = false)
        {
            var fullPath = tempFile
                ? Path.Combine(folder, Path.GetRandomFileName())
                : Path.Combine(folder, Name + Type.FileExtension());
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

        public IVBProject ParentProject { get; private set; }

        private void ExportUserFormModule(string path)
        {
            // VBIDE API inserts an extra newline when exporting a UserForm module.
            // this issue causes forms to always be treated as "modified" in source control, which causes conflicts.
            // we need to remove the extra newline before the file gets written to its output location.

            int legitEmptyLineCount;
            using (var codeModule = CodeModule)
            {
                var visibleCode = codeModule.Content().Split(new[] {Environment.NewLine}, StringSplitOptions.None);
                legitEmptyLineCount = visibleCode.TakeWhile(string.IsNullOrWhiteSpace).Count();
            }

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
            using (var codeModule = CodeModule)
            {
                var lineCount = codeModule.CountOfLines;
                if (lineCount > 0)
                {
                    var text = codeModule.GetLines(1, lineCount);
                    File.WriteAllText(path, text);
                }
            }
        }

        private string ExportToTempFile()
        {
            var path = Path.Combine(Path.GetTempPath(), Name + Type.FileExtension());
            Export(path);
            return path;
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
    }
}