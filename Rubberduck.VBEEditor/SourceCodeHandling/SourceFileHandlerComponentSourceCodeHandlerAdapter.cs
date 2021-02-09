using System.IO.Abstractions;
using System.Text;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public class SourceFileHandlerComponentSourceCodeHandlerAdapter : IComponentSourceCodeHandler
    {
        private readonly ITempSourceFileHandler _tempSourceFileHandler;
        private readonly IFileSystem _fileSystem;

        public SourceFileHandlerComponentSourceCodeHandlerAdapter(
            ITempSourceFileHandler tempSourceFileHandler,
            IFileSystem fileSystem)
        {
            _tempSourceFileHandler = tempSourceFileHandler;
            _fileSystem = fileSystem;
        }

        public string SourceCode(IVBComponent module)
        {
            return _tempSourceFileHandler.Read(module) ?? string.Empty;
        }

        public IVBComponent SubstituteCode(IVBComponent module, string newCode)
        {
            if (module.Type == ComponentType.Document)
            {
                //We cannot substitute the code of a document module via the file.
                return module;
            }

            var fileName = _tempSourceFileHandler.Export(module);
            if (fileName == null || !_fileSystem.File.Exists(fileName))
            {
                return module;
            }
            _fileSystem.File.WriteAllText(fileName, newCode, Encoding.Default);
            return _tempSourceFileHandler.ImportAndCleanUp(module, fileName);
        }
    }
}
