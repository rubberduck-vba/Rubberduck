using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class FolderAnnotation : AnnotationBase
    {
        private readonly string _folderName;

        public FolderAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.Folder, qualifiedSelection)
        {
            _folderName = parameters.FirstOrDefault() ?? string.Empty;
        }

        public string FolderName
        {
            get
            {
                return _folderName;
            }
        }

        public override string ToString()
        {
            return string.Format("Folder: {0}", _folderName);
        }
    }
}
