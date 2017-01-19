using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class UntypedFunctionUsageInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private IEnumerable<QuickFixBase> _quickFixes;

        public UntypedFunctionUsageInspectionResult(IInspection inspection, IdentifierReference reference) 
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new UntypedFunctionUsageQuickFix((ParserRuleContext)GetFirst(typeof(VBAParser.IdentifierContext)).Parent, QualifiedSelection), 
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.UntypedFunctionUsageInspectionResultFormat, _reference.Declaration.IdentifierName).Captialize(); }
        }

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
