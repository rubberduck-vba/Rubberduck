using System.IO.Abstractions;
using System.Text;
using Rubberduck.InternalApi.Common;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.VBEditor.VB6
{
    public class ExternalFileTempSourceFileHandlerEmulator : ITempSourceFileHandler
    {
        private IFileSystem _fileSystem => FileSystemProvider.FileSystem;

        public string Export(IVBComponent component)
        {
            // VB6 source code is already external, and should be in the first associated file.
            return component.GetFileName(1);            
        }

        public IVBComponent ImportAndCleanUp(IVBComponent component, string fileName)
        {
            // VB6 source code can be written directly in-place, without needing to import it, hence no-op.
            return component;
        }

        public string Read(IVBComponent component)
        {
            var fileName = Export(component);
            if (fileName == null || !_fileSystem.File.Exists(fileName))
            {
                return null;
            }

            return _fileSystem.File.ReadAllText(fileName, Encoding.Default);
        }        
    }
}
