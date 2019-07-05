using System.IO;
using System.Text;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.VBEditor.VB6
{
    public class ExternalFileTempSourceFileHandlerEmulator : ITempSourceFileHandler
    {
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
            if (fileName == null || !File.Exists(fileName))
            {
                return null;
            }

            return File.ReadAllText(fileName, Encoding.Default);
        }        
    }
}
