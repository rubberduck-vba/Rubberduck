using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class UntypedFunctionUsageInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;

        public UntypedFunctionUsageInspectionResult(IInspection inspection, IdentifierReference reference, QualifiedMemberName? qualifiedName) 
            : base(inspection, reference.QualifiedModuleName, qualifiedName, reference.Context)
        {
            _reference = reference;
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.UntypedFunctionUsageInspectionResultFormat, _reference.Declaration.IdentifierName).Capitalize(); }
        }

        // note: remove before PRing
        private ParserRuleContext GetFirst(Type nodeType)
        {
            var unexploredNodes = new List<ParserRuleContext> {Context};

            while (unexploredNodes.Any())
            {
                if (unexploredNodes[0].GetType() == nodeType)
                {
                    return unexploredNodes[0];
                }
                
                unexploredNodes.AddRange(unexploredNodes[0].children.OfType<ParserRuleContext>());
                unexploredNodes.RemoveAt(0);
            }

            return null;
        }
    }
}
