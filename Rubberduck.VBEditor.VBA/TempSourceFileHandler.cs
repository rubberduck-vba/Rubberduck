using System.IO.Abstractions;
using System.Text;
using Rubberduck.InternalApi.Common;
using Rubberduck.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.VBEditor.VBA
{
    public class TempSourceFileHandler : ITempSourceFileHandler
    {
        private IFileSystem _fileSystem = FileSystemProvider.FileSystem;

        public string Export(IVBComponent component)
        {
            if (!_fileSystem.Directory.Exists(ApplicationConstants.RUBBERDUCK_TEMP_PATH))
            {
                _fileSystem.Directory.CreateDirectory(ApplicationConstants.RUBBERDUCK_TEMP_PATH);
            }
            var fileName = component.ExportAsSourceFile(ApplicationConstants.RUBBERDUCK_TEMP_PATH, true, false);

            return _fileSystem.File.Exists(fileName) 
                ? fileName 
                : null;         
        }

        public IVBComponent ImportAndCleanUp(IVBComponent component, string fileName)
        {
            if (fileName == null || !_fileSystem.File.Exists(fileName))
            {
                return component;
            }

            IVBComponent newComponent = null;
            using (var components = component.Collection)
            {
                components.Remove(component);
                newComponent = components.ImportSourceFile(fileName);
            }

            try
            {
                _fileSystem.File.Delete(fileName);
            }
            catch
            {
                // Meh.
            }

            return newComponent;
        }

        public string Read(IVBComponent component)
        {
            var fileName = Export(component);
            if (fileName == null || !_fileSystem.File.Exists(fileName))
            {
                return null;
            }

            var code = _fileSystem.File.ReadAllText(fileName, Encoding.Default);
            try
            {
                _fileSystem.File.Delete(fileName);
            }
            catch
            {
                // Meh.
            }            

            return code;
        }
        
    }
}
