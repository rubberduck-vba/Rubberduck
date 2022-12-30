using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.SmartIndenter
{

    public interface IIndenter : ISimpleIndenter
    {
        void IndentCurrentProcedure();
        void IndentCurrentModule();
        void IndentCurrentProject();
        void Indent(IVBComponent component);
    }
}
