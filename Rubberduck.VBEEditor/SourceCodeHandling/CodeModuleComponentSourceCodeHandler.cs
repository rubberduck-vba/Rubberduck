using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public class CodeModuleComponentSourceCodeHandler : IComponentSourceCodeHandler
    {
        public string SourceCode(IVBComponent module)
        {
            using (var codeModule = module.CodeModule)
            {
                return codeModule.Content() ?? string.Empty;
            }
        }

        public IVBComponent SubstituteCode(IVBComponent module, string newCode)
        {
            using (var codeModule = module.CodeModule)
            {
                codeModule.Clear();
                codeModule.InsertLines(1, newCode);
            }

            return module;
        }
    }
}