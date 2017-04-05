namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IQuickFix
    {
        void Fix(IInspectionResult result);
        string Description(IInspectionResult result);

        bool CanFixInProcedure { get; }
        bool CanFixInModule { get; }
        bool CanFixInProject { get; }
    }
}