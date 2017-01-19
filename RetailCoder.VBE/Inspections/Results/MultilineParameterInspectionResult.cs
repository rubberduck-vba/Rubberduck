using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class MultilineParameterInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;

        public MultilineParameterInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        { }

        public override string Description
        {
            get
            {
                return string.Format(
                    Target.Context.GetSelection().LineCount > 3
                        ? RubberduckUI.EasterEgg_Continuator
                        : InspectionsUI.MultilineParameterInspectionResultFormat, Target.IdentifierName).Captialize();
            }
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new MakeSingleLineParameterQuickFix(Context, QualifiedSelection),
                    new IgnoreOnceQuickFix(Target.ParentDeclaration.Context, Target.ParentDeclaration.QualifiedSelection, Inspection.AnnotationName) 
                });
            }
        }
    }
}
