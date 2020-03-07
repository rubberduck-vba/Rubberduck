using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.CodeAnalysis.Inspections
{
    public interface IParseTreeInspection : IInspection
    {
        CodeKind TargetKindOfCode { get; }
        IInspectionListener Listener { get; }
    }
}
