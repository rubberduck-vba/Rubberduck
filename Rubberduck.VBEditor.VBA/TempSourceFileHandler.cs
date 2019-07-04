using System.IO;
using System.Text;
using Rubberduck.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.VBEditor.VBA
{
    public class TempSourceFileHandler : ITempSourceFileHandler
    {
        public string Export(IVBComponent component)
        {
            if (!Directory.Exists(ApplicationConstants.RUBBERDUCK_TEMP_PATH))
            {
                Directory.CreateDirectory(ApplicationConstants.RUBBERDUCK_TEMP_PATH);
            }
            var fileName = component.ExportAsSourceFile(ApplicationConstants.RUBBERDUCK_TEMP_PATH, true, false);

            return File.Exists(fileName) 
                ? fileName 
                : null;         
        }

        public IVBComponent ImportAndCleanUp(IVBComponent component, string fileName)
        {
            if (fileName == null || !File.Exists(fileName))
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
                File.Delete(fileName);
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
            if (fileName == null || !File.Exists(fileName))
            {
                return null;
            }

            var code = File.ReadAllText(fileName, Encoding.Default);
            try
            {
                File.Delete(fileName);
            }
            catch
            {
                // Meh.
            }            

            return code;
        }
        
    }
}
