using System.Collections.Generic;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IParseTreeInspection : IInspection
    {
        void SetResults(IEnumerable<QualifiedContext> results);
    }
}
