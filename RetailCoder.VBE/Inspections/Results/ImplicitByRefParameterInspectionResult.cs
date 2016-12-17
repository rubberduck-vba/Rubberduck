using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitByRefParameterInspectionResult : InspectionResultBase
    {
        private readonly string _identifierName;
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ImplicitByRefParameterInspectionResult(IInspection inspection, string identifierName, QualifiedContext<VBAParser.ArgContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _identifierName = identifierName;
            _quickFixes = new QuickFixBase[]
                {
                    new ChangeParameterByRefByValQuickFix(Context, QualifiedSelection, InspectionsUI.ImplicitByRefParameterQuickFix, Tokens.ByRef), 
                    new IgnoreOnceQuickFix(qualifiedContext.Context, QualifiedSelection, Inspection.AnnotationName), 
                };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ImplicitByRefParameterInspectionResultFormat, _identifierName).Captialize(); }
        }
    }
}
