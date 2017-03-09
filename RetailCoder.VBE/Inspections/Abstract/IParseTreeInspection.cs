using System.Collections.Generic;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Inspections.Abstract
{
    public interface IParseTreeInspection : IInspection
    {
        void SetResults(IEnumerable<QualifiedContext> results);
    }
}
