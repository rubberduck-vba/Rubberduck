using System;
using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteByRefModifierInspectionResult : InspectionResultBase
    {
        private readonly Lazy<IEnumerable<QuickFixBase>> _quickFixes;

        public ObsoleteByRefModifierInspectionResult(IInspection inspection, Declaration declaration)
            : base(inspection, declaration)
        {
            _quickFixes = new Lazy<IEnumerable<QuickFixBase>>(() =>
                new QuickFixBase[]
                {
                    new RemoveExplicitByRefModifierQuickFix(Context, QualifiedSelection),
                    new IgnoreOnceQuickFix(Target.Context, QualifiedSelection, Inspection.AnnotationName)
                });
        }

        public ObsoleteByRefModifierInspectionResult(IInspection inspection, Declaration interfaceDeclaration, IEnumerable<Declaration> implementationDeclarations)
            : base(inspection, interfaceDeclaration)
        {
            _quickFixes = new Lazy<IEnumerable<QuickFixBase>>(() =>
                new QuickFixBase[]
                {
                    new RemoveExplicitByRefModifierQuickFix(Context, QualifiedSelection)
                    {
                       InterfaceImplementationDeclarations = implementationDeclarations 
                    },
                    new IgnoreOnceQuickFix(Target.Context, QualifiedSelection, Inspection.AnnotationName)
                });
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes.Value; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ObsoleteByRefModifierInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }
    }
}
