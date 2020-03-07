using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about Rubberduck annotations with more arguments than allowed.
    /// </summary>
    /// <why>
    /// Most annotations only process a limited number of arguments; superfluous arguments are ignored.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// '@Folder "MyFolder.MySubFolder" "SomethingElse
    /// '@PredeclaredId True
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// '@Folder("MyFolder.MySubFolder")
    /// '@PredeclaredId
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    internal sealed class SuperfluousAnnotationArgumentInspection : InspectionBase
    {
        public SuperfluousAnnotationArgumentInspection(IDeclarationFinderProvider declarationFinderProvider)
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
            var allowedArguments = pta.Annotation.AllowedArguments;
            return allowedArguments.HasValue && allowedArguments.Value < pta.AnnotationArguments.Count;
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
                InspectionResults.SuperfluousAnnotationArgumentInspection,
                pta.Annotation.Name);
        }
    }
}
