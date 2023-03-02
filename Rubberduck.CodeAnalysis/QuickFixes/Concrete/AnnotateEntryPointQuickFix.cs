using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adds an '@EntryPoint annotation to mark as an entry point a procedure that isn't intended to be invoked by any code.
    /// </summary>
    /// <inspections>
    /// <inspection name="ProcedureNotUsedInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// '@EntryPoint
    /// Public Sub DoSomething()
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class AnnotateEntryPointQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;
        private readonly IAnnotationUpdater _annotationUpdater;

        public AnnotateEntryPointQuickFix(IAnnotationUpdater annotationUpdater, RubberduckParserState state)
            : base(new[] { typeof(Inspections.Concrete.ProcedureNotUsedInspection) }.ToArray())
        {
            _state = state;
            _annotationUpdater = annotationUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var module = result.QualifiedSelection.QualifiedName;
            var lineToAnnotate = result.QualifiedSelection.Selection.StartLine;
            var existingEntryPointAnnotation = _state.DeclarationFinder
                .FindAnnotations(module, lineToAnnotate)
                .FirstOrDefault(pta => pta.Annotation is EntryPointAnnotation);

            if (existingEntryPointAnnotation == null)
            {
                _annotationUpdater.AddAnnotation(rewriteSession, new QualifiedContext(module, result.Context), new EntryPointAnnotation());
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AnnotateEntryPointQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
