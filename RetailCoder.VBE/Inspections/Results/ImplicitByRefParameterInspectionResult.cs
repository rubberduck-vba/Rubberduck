using System;
using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitByRefParameterInspectionResult : InspectionResultBase
    {
        private readonly Lazy<IEnumerable<IQuickFix>> _quickFixes;

        public ImplicitByRefParameterInspectionResult(IInspection inspection, Declaration declaration)
            : base(inspection, declaration)
        {
            _quickFixes = new Lazy<IEnumerable<IQuickFix>>(() =>
                new IQuickFix[]
                {
                    new ChangeParameterByRefByValQuickFix(Context, QualifiedSelection, InspectionsUI.ImplicitByRefParameterQuickFix, Tokens.ByRef),
                    new IgnoreOnceQuickFix(Target.Context, QualifiedSelection, Inspection.AnnotationName)
                });
        }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get { return _quickFixes.Value; }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ImplicitByRefParameterInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }
    }
}
