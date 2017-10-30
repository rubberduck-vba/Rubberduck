using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IParseTreeInspection : IInspection
    {
        ParsePass Pass { get; }
        IInspectionListener Listener { get; }
    }
}
