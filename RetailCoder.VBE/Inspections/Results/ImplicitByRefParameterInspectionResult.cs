using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitByRefParameterInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ImplicitByRefParameterInspectionResult(IInspection inspection, Declaration declaration)
            : base(inspection, declaration)
        {
            _quickFixes = new QuickFixBase[]
                {
                    new ChangeParameterByRefByValQuickFix(Context, declaration.QualifiedSelection, InspectionsUI.ImplicitByRefParameterQuickFix, Tokens.ByRef), 
                    new IgnoreOnceQuickFix(declaration.Context, declaration.QualifiedSelection, Inspection.AnnotationName), 
                };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ImplicitByRefParameterInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }
    }
}
