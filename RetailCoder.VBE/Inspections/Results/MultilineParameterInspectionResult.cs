using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class MultilineParameterInspectionResult : InspectionResultBase
    {
        private IEnumerable<IQuickFix> _quickFixes;

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

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new IQuickFix[]
                {
                    new MakeSingleLineParameterQuickFix(Context, QualifiedSelection),
                    new IgnoreOnceQuickFix(Target.ParentDeclaration.Context, Target.ParentDeclaration.QualifiedSelection, Inspection.AnnotationName) 
                });
            }
        }
    }
}
