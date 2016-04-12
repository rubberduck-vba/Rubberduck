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
            if (parameters.Count() != 1)
            {
                throw new InvalidAnnotationArgumentException(string.Format("{0} expects exactly one argument, the folder, but none or more than one were passed.", this.GetType().Name));
            }
            _folderName = parameters.First();
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
