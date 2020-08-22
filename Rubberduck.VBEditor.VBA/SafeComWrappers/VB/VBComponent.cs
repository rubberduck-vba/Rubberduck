using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBComponent : SafeComWrapper<VB.VBComponent>, IVBComponent
    {
        public VBComponent(VB.VBComponent target, bool rewrapping = false) 
            : base(target, rewrapping)
        { }

        public QualifiedModuleName QualifiedModuleName => new QualifiedModuleName(this);

        public ComponentType Type => IsWrappingNullReference ? 0 : (ComponentType)Target.Type;
        public bool HasCodeModule => true;
        public ICodeModule CodeModule => new CodeModule(IsWrappingNullReference ? null : Target.CodeModule);
        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);
        public IVBComponents Collection => new VBComponents(IsWrappingNullReference ? null : Target.Collection);
        public IProperties Properties => new Properties(IsWrappingNullReference ? null : Target.Properties);
        public bool HasOpenDesigner => !IsWrappingNullReference && Target.HasOpenDesigner;
        public string DesignerId => IsWrappingNullReference ? string.Empty : Target.DesignerID;

        public string Name
        {
            get => IsWrappingNullReference ? string.Empty : Target.Name;
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
                    : new UserForm(Target.Designer as VB.Forms.UserForm))
                {
                    return designer == null
                        ? new Controls(null)
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
                    : new UserForm(Target.Designer as VB.Forms.UserForm))
                {
                    return designer == null
                        ? new Controls(null)
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
                using (var designer = new UserForm(Target.Designer as VB.Forms.UserForm))
                {
                    return !designer.IsWrappingNullReference;
                }
            }
        }

        public IWindow DesignerWindow() => new Window(IsWrappingNullReference ? null : Target.DesignerWindow());
        public void Activate() => Target.Activate();
        public bool IsSaved => !IsWrappingNullReference && Target.Saved;
        public void Export(string path) => Target.Export(path);

        /// <summary>
        /// Exports the component to the folder. The file name matches the component name and file extension is based on the component's type.
        /// </summary>
        /// <param name="folder">Destination folder for the resulting source file.</param>
        /// <param name="isTempFile">True if a unique temp file name should be generated. WARNING: filenames generated with this flag are not persisted.</param>
        /// <param name="specialCaseDocumentModules">If reimport of a document file is required later, it has to receive special treatment.</param>
        public string ExportAsSourceFile(string folder, bool isTempFile = false, bool specialCaseDocumentModules = true)
        {
            //TODO: this entire thign needs to be reworked. IO is not the class' concern.
            //We probably need to leverage IPersistancePathProvider? ITempSourceFileHandler? 
            //Just not here.
            var fullPath = isTempFile
                ? Path.Combine(folder, Path.GetRandomFileName())
                : Path.Combine(folder, SafeName + Type.FileExtension());

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            switch (Type)
            {
                case ComponentType.UserForm:
                    ExportUserFormModule(fullPath);
                    break;
                case ComponentType.Document:
                    if(specialCaseDocumentModules)
                    {
                        ExportDocumentModule(fullPath);
                    }
                    else
                    {
                        Export(fullPath);
                    }
                    break;
                default:
                    Export(fullPath);
                    break;
            }

            return fullPath;
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

        private void ExportUserFormModule(string path)
        {
            // VBIDE API inserts an extra newline when exporting a UserForm module.
            // this issue causes forms to always be treated as "modified" in source control, which causes conflicts.
            // we need to remove the extra newline before the file gets written to its output location.

            int legitEmptyLineCount;
            using (var codeModule = CodeModule)
            {
                var visibleCode = codeModule.Content().Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                legitEmptyLineCount = visibleCode.TakeWhile(string.IsNullOrWhiteSpace).Count();
            }

            var tempFile = ExportToTempFile();
            var tempFilePath = Directory.GetParent(tempFile).FullName;
            var fileEncoding = Encoding.Default;    //We use the current ANSI codepage because that is what the VBE does.
            var contents = File.ReadAllLines(tempFile, fileEncoding);
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
            File.WriteAllLines(path, code, fileEncoding);

            // LINQ hates this search, therefore, iterate the long way
            foreach (string line in contents)
            {
                if (line.Contains("OleObjectBlob"))
                {
                    var binaryFileName = line.Trim().Split('"')[1];
                    var destPath = Directory.GetParent(path).FullName;
                    if (File.Exists(Path.Combine(tempFilePath, binaryFileName)) && !destPath.Equals(tempFilePath))
                    {
                        System.Diagnostics.Debug.WriteLine(Path.Combine(destPath, binaryFileName));
                        if (File.Exists(Path.Combine(destPath, binaryFileName)))
                        {
                            try
                            {
                                File.Delete(Path.Combine(destPath, binaryFileName));
                            }
                            catch (Exception)
                            {
                                // Meh?
                            }
                        }
                        File.Copy(Path.Combine(tempFilePath, binaryFileName), Path.Combine(destPath, binaryFileName));
                    }
                    break;
                }
            }
        }

        private void ExportDocumentModule(string path)
        {
            using (var codeModule = CodeModule)
            {
                var lineCount = codeModule.CountOfLines;
                if (lineCount > 0)
                {
                    //One cannot reimport document modules as such in the VBE; so we simply export and import the contents of the code pane.
                    //Because of this, it is OK, and actually preferable, to use the default UTF8 encoding.
                    var text = codeModule.GetLines(1, lineCount);
                    File.WriteAllText(path, text, Encoding.UTF8);
                }
            }
        }

        private string ExportToTempFile()
        {
            var path = Path.Combine(Path.GetTempPath(), SafeName + Type.FileExtension());
            Export(path);
            return path;
        }

        public override bool Equals(ISafeComWrapper<VB.VBComponent> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public int FileCount => 0;
        public string GetFileName(short index) => null;

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