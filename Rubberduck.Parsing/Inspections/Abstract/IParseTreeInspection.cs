using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IParseTreeInspection : IInspection
    {
        CodeKind TargetKindOfCode { get; }
        IInspectionListener Listener { get; }
    }
}
