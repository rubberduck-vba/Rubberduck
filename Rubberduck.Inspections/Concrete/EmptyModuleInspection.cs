using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    class EmptyModuleInspection : InspectionBase
    {
        public EmptyModuleInspection(RubberduckParserState state, CodeInspectionSeverity defaultSeverity = CodeInspectionSeverity.Hint) 
        :base(state, defaultSeverity)
        { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            throw new NotImplementedException();
        }
    }
}
