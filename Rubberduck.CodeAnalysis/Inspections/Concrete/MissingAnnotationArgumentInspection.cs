using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about a malformed Rubberduck annotation that is missing one or more arguments.
    /// </summary>
    /// <why>
    /// Some annotations require arguments; if the required number of arguments isn't specified, the annotation is nothing more than an obscure comment.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Folder
    /// '@ModuleDescription
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Folder("MyProject.XYZ")
    /// '@ModuleDescription("This module does XYZ")
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class MissingAnnotationArgumentInspection : InspectionBase
    {
        public MissingAnnotationArgumentInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            return finder.UserDeclarations(DeclarationType.Module)
                .Where(module => module != null)
                .SelectMany(module => DoGetInspectionResults(module.QualifiedModuleName, finder))
                .ToList();
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableAnnotations = finder.FindAnnotations(module)
                .Where(IsResultAnnotation);

            return objectionableAnnotations
                .Select(InspectionResult)
                .ToList();
        }

        private static bool IsResultAnnotation(IParseTreeAnnotation pta)
        {
            return pta.Annotation.RequiredArguments > pta.AnnotationArguments.Count;
        }

        private IInspectionResult InspectionResult(IParseTreeAnnotation pta)
        {
            var qualifiedContext = new QualifiedContext(pta.QualifiedSelection.QualifiedName, pta.Context);
            return new QualifiedContextInspectionResult(
                this,
                ResultDescription(pta),
                qualifiedContext);
        }

        private static string ResultDescription(IParseTreeAnnotation pta)
        {
            return string.Format(
                InspectionResults.MissingAnnotationArgumentInspection,
                pta.Annotation.Name);
        }
    }
}
