using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IParseTreeInspection : IInspection
    {
        CodeKind TargetKindOfCode { get; }
        IInspectionListener Listener { get; }
    }
}
