using System.IO;
using System.Text;
using Rubberduck.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.VBA
{
    public class SourceCodeHandler : ISourceCodeHandler
    {
        public string Export(IVBComponent component)
        {
            var fileName = component.ExportAsSourceFile(ApplicationConstants.RUBBERDUCK_TEMP_PATH);

            return File.Exists(fileName) 
                ? fileName 
                : null; // a document component without any code wouldn't be exported (file would be empty anyway).            
        }

        public void Import(IVBComponent component, string fileName)
        {
            using (var components = component.Collection)
            {
                components.Remove(component);
                components.ImportSourceFile(fileName);
            }
        }

        public string Read(IVBComponent component)
        {
            var fileName = Export(component);
            if (fileName == null)
            {
                return null;
            }

            var encoding = component.QualifiedModuleName.ComponentType == ComponentType.Document
                ? Encoding.UTF8
                : Encoding.Default;

            var code = File.ReadAllText(fileName, encoding);
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
