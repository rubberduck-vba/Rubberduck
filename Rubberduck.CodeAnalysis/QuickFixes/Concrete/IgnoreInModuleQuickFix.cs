using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adds an '@IgnoreModule annotation to ignore a inspection results for a specific inspection inside a whole module. Applicable to all inspections whose results can be annotated in a module.
    /// </summary>
    /// <canfix multiple="true" procedure="false" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     Dim bar As Long
    ///     value = 42
    ///     bar = 23
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// '@IgnoreModule VariableNotUsed
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     Dim bar As Long
    ///     value = 42
    ///     bar = 23
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class IgnoreInModuleQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;
        private readonly IAnnotationUpdater _annotationUpdater;

        public IgnoreInModuleQuickFix(IAnnotationUpdater annotationUpdater, RubberduckParserState state, IEnumerable<IInspection> inspections)
            : base(inspections.Select(s => s.GetType()).Where(i => i.CustomAttributes.All(a => a.AttributeType != typeof(CannotAnnotateAttribute))).ToArray())
        {
            _state = state;
            _annotationUpdater = annotationUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var module = result.QualifiedSelection.QualifiedName;
            var moduleDeclaration = _state.DeclarationFinder.Members(module, DeclarationType.Module)
                .FirstOrDefault();

            if (moduleDeclaration == null)
            {
                return;
            }

            var existingIgnoreModuleAnnotation = moduleDeclaration.Annotations
                .FirstOrDefault(pta => pta.Annotation is IgnoreModuleAnnotation);

            var annotationType = new IgnoreModuleAnnotation();
            if (existingIgnoreModuleAnnotation != null)
            {
                var annotationValues = existingIgnoreModuleAnnotation.AnnotationArguments.ToList();

                if (annotationValues.Contains(result.Inspection.AnnotationName))
                {
                    return;
                }

                annotationValues.Insert(0, result.Inspection.AnnotationName);
                _annotationUpdater.UpdateAnnotation(rewriteSession, existingIgnoreModuleAnnotation, annotationType, annotationValues);
            }
            else
            {
                var newModuleText = rewriteSession.CheckOutModuleRewriter(module).GetText();
                var ignoreModuleText = $"'{ParseTreeAnnotation.ANNOTATION_MARKER}{annotationType.Name}";
                if (newModuleText.Contains(ignoreModuleText))
                {
                    //Most probably, we have added this already in another invocation on the same rewrite session. 
                    return;
                }

                var annotationValues = new List<string> { result.Inspection.AnnotationName };
                _annotationUpdater.AddAnnotation(rewriteSession, moduleDeclaration, annotationType, annotationValues);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IgnoreInModuleQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}