using System.IO;
using System.Text;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.VB6
{
    public class SourceCodeHandler : ISourceCodeHandler
    {
        public string Export(IVBComponent component)
        {
            // VB6 source code is already external, and should be in the first associated file.
            return component.GetFileName(1);            
        }

        public void Import(IVBComponent component, string fileName)
        {
            // VB6 source code can be written directly in-place, without needing to import it, hence no-op.
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

            return File.ReadAllText(fileName, encoding);
        }        
    }
}
