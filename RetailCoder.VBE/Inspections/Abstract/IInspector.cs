using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Abstract
{
    public interface IInspector : IDisposable
    {
        Task<IEnumerable<IInspectionResult>> FindIssuesAsync(RubberduckParserState state, CancellationToken token);
    }
}
