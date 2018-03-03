using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.Navigation.CodeMetrics
{
    public interface ICodeMetricsAnalyst
    {
        IEnumerable<ModuleMetricsResult> ModuleMetrics(RubberduckParserState state);
    }
}
