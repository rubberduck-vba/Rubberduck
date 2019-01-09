using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class IgnoreOnceQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;
        private readonly IAnnotationUpdater _annotationUpdater;

        public IgnoreOnceQuickFix(IAnnotationUpdater annotationUpdater, RubberduckParserState state, IEnumerable<IInspection> inspections)
            : base(inspections.Select(s => s.GetType()).Where(i => i.CustomAttributes.All(a => a.AttributeType != typeof(CannotAnnotateAttribute))).ToArray())
        {
            _state = state;
            _annotationUpdater = annotationUpdater;
        }

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (result.Target?.DeclarationType.HasFlag(DeclarationType.Module) ?? false)
            {
                FixModule(result, rewriteSession);
            }
            else
            {
                FixNonModule(result, rewriteSession);
            }
        }

        private void FixNonModule(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var module = result.QualifiedSelection.QualifiedName;
            var lineToAnnotate = result.QualifiedSelection.Selection.StartLine;
            var existingIgnoreAnnotation = _state.DeclarationFinder.FindAnnotations(module, lineToAnnotate)
                .OfType<IgnoreAnnotation>()
                .FirstOrDefault();

            var annotationType = AnnotationType.Ignore;
            if (existingIgnoreAnnotation != null)
            {
                var annotationValues = existingIgnoreAnnotation.InspectionNames.ToList();
                annotationValues.Insert(0, result.Inspection.AnnotationName);
                _annotationUpdater.UpdateAnnotation(rewriteSession, existingIgnoreAnnotation, annotationType, annotationValues);
            }
            else
            {
                var annotationValues = new List<string> { result.Inspection.AnnotationName };
                _annotationUpdater.AddAnnotation(rewriteSession, new QualifiedContext(module, result.Context), annotationType, annotationValues);
            }
        }

        private void FixModule(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var moduleDeclaration = result.Target;
            var existingIgnoreModuleAnnotation = moduleDeclaration.Annotations
                .OfType<IgnoreModuleAnnotation>()
                .FirstOrDefault();

            var annotationType = AnnotationType.IgnoreModule;
            if (existingIgnoreModuleAnnotation != null)
            {
                var annotationValues = existingIgnoreModuleAnnotation.InspectionNames.ToList();
                annotationValues.Insert(0, result.Inspection.AnnotationName);
                _annotationUpdater.UpdateAnnotation(rewriteSession, existingIgnoreModuleAnnotation, annotationType, annotationValues);
            }
            else
            {
                var annotationValues = new List<string> { result.Inspection.AnnotationName };
                _annotationUpdater.AddAnnotation(rewriteSession, moduleDeclaration, annotationType, annotationValues);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IgnoreOnce;
    }
}
