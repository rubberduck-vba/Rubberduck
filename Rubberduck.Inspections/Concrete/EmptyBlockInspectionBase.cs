using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete
{
    public class EmptyBlockInspectionBase<T> : ParseTreeInspectionBase
    {
        private readonly string _resultFormat;

        public EmptyBlockInspectionBase(RubberduckParserState state, string resultFormat) : base(state, CodeInspectionSeverity.Suggestion)
        {
            _resultFormat = resultFormat;
        }

        public override Type Type => typeof(T);

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IInspectionListener Listener => throw new NotImplementedException();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        _resultFormat,
                                                        result));
        }
    }
}
