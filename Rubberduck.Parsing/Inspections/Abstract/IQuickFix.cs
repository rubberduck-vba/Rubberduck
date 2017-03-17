using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IQuickFix
    {
        void Fix();
        bool CanFixInProject { get; }
        bool CanFixInModule { get; }
        bool CanFixInProcedure { get; }

        string Description { get; }
        bool IsCancelled { get; set; }
        QualifiedSelection Selection { get; }
    }
}