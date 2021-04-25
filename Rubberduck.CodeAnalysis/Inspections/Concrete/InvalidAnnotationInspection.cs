using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.InternalApi.Extensions;
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
    /// An inspection that flags invalid annotation comments.
    /// </summary>
    internal abstract class InvalidAnnotationInspectionBase : InspectionBase
    {
        protected InvalidAnnotationInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        protected QualifiedContext Context(IParseTreeAnnotation pta) =>
            new QualifiedContext(pta.QualifiedSelection.QualifiedName, pta.Context);

        protected sealed override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            return finder.UserDeclarations(DeclarationType.Module)
.Where(module => module != null)
.SelectMany(module => DoGetInspectionResults(module.QualifiedModuleName, finder));
        }

        protected IInspectionResult InspectionResult(IParseTreeAnnotation pta) =>
            new QualifiedContextInspectionResult(this, ResultDescription(pta), Context(pta));

        /// <summary>
        /// Gets all invalid annotations covered by this inspection.
        /// </summary>
        /// <param name="annotations">All user code annotations.</param>
        /// <param name="userDeclarations">All user declarations.</param>
        /// <param name="identifierReferences">All identifier references in user code.</param>
        /// <returns></returns>
        protected abstract IEnumerable<IParseTreeAnnotation> GetInvalidAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations,
            IEnumerable<Declaration> userDeclarations,
            IEnumerable<IdentifierReference> identifierReferences);

        /// <summary>
        /// Gets an annotation-specific description for an inspection result.
        /// </summary>
        /// <param name="pta">The invalid annotation.</param>
        /// <returns></returns>
        protected abstract string ResultDescription(IParseTreeAnnotation pta);

        protected sealed override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var annotations = finder.FindAnnotations(module);
            var userDeclarations = finder.Members(module).ToList();
            var identifierReferences = finder.IdentifierReferences(module).ToList();

            var invalidAnnotations = GetInvalidAnnotations(annotations, userDeclarations, identifierReferences);
            return invalidAnnotations.Select(InspectionResult).ToList();
        }
    }

    /// <summary>
    /// Flags comments that parsed like Rubberduck annotations, but were not recognized as such.
    /// </summary>
    /// <why>
    /// Other add-ins may support similar-looking annotations that Rubberduck does not recognize; this inspection can be used to spot a typo in Rubberduck annotations.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Param "Value", "The value to print."  : Rubberduck does not define a @Param annotation
    /// Public Sub Test(ByVal Value As Long)
    ///     Debug.Print Value
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// '@Description "Prints the specified value." : Rubberduck defines a @Description annotation
    /// Public Sub Test(ByVal Value As Long)
    ///     Debug.Print Value
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UnrecognizedAnnotationInspection : InvalidAnnotationInspectionBase
    {
        public UnrecognizedAnnotationInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        protected override IEnumerable<IParseTreeAnnotation> GetInvalidAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations, 
            IEnumerable<Declaration> userDeclarations, 
            IEnumerable<IdentifierReference> identifierReferences)
        {
            return annotations.Where(pta => pta.Annotation is NotRecognizedAnnotation).ToList();
        }

        protected override string ResultDescription(IParseTreeAnnotation pta) =>
            string.Format(InspectionResults.UnrecognizedAnnotationInspection, pta.Context.GetText());
    }

    /// <summary>
    /// Flags Rubberduck annotations used in a component type that is incompatible with that annotation.
    /// </summary>
    /// <why>
    /// Some annotations can only be used in a specific type of module; others cannot be used in certain types of modules.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@PredeclaredId  'this annotation is illegal in a standard module
    /// Option Explicit
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// '@PredeclaredId  'this annotation works fine in a class module
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class AnnotationInIncompatibleComponentTypeInspection : InvalidAnnotationInspectionBase
    {
        public AnnotationInIncompatibleComponentTypeInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        protected override IEnumerable<IParseTreeAnnotation> GetInvalidAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations,
            IEnumerable<Declaration> userDeclarations,
            IEnumerable<IdentifierReference> identifierReferences)
        {
            foreach (var pta in annotations)
            {
                var annotation = pta.Annotation;
                var componentType = pta.QualifiedSelection.QualifiedName.ComponentType;
                if (annotation.RequiredComponentType.HasValue && annotation.RequiredComponentType != componentType
                       || annotation.IncompatibleComponentTypes.Contains(componentType))
                {
                    yield return pta;
                }
            }

            yield break;
        }

        protected override string ResultDescription(IParseTreeAnnotation pta)
        {
            if (pta.Annotation.RequiredComponentType.HasValue)
            {
                return string.Format(InspectionResults.InvalidAnnotationInspection_NotInRequiredComponentType,
                    pta.Annotation.Name, // annotation...
                    pta.QualifiedSelection.QualifiedName.ComponentType,  // is used in a...
                    pta.Annotation.RequiredComponentType); // but is only valid in a...
            }
            else
            {
                return string.Format(InspectionResults.InvalidAnnotationInspection_IncompatibleComponentType,
                    pta.Annotation.Name, // annotation...
                    pta.QualifiedSelection.QualifiedName.ComponentType); // cannot be used in a...
            }
        }
    }

    /// <summary>
    /// Flags invalid or misplaced Rubberduck annotation comments.
    /// </summary>
    /// <why>
    /// Rubberduck is correctly parsing an annotation, but that annotation is illegal in that context and couldn't be bound to a code element.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     '@Folder("Module1.DoSomething")
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Folder("Module1.DoSomething")
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class InvalidAnnotationInspection : InvalidAnnotationInspectionBase
    {
        public InvalidAnnotationInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override string ResultDescription(IParseTreeAnnotation pta) =>
            string.Format(InspectionResults.InvalidAnnotationInspection, pta.Annotation.Name);

        protected override IEnumerable<IParseTreeAnnotation> GetInvalidAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations,
            IEnumerable<Declaration> userDeclarations,
            IEnumerable<IdentifierReference> identifierReferences)
        {
            return GetUnboundAnnotations(annotations, userDeclarations, identifierReferences)
                .Where(pta => !pta.Annotation.Target.HasFlag(AnnotationTarget.General) || pta.AnnotatedLine == null)
                .Concat(AttributeAnnotationsOnDeclarationsNotAllowingAttributes(userDeclarations))
                .ToList();
        }

        private IEnumerable<IParseTreeAnnotation> GetUnboundAnnotations(
            IEnumerable<IParseTreeAnnotation> annotations,
            IEnumerable<Declaration> userDeclarations,
            IEnumerable<IdentifierReference> identifierReferences)
        {
            var boundAnnotationsSelections = userDeclarations
                .SelectMany(declaration => declaration.Annotations)
                .Concat(identifierReferences.SelectMany(reference => reference.Annotations))
                .Select(annotation => annotation.QualifiedSelection)
                .ToHashSet();

            return annotations
                .Where(pta => pta.Annotation.GetType() != typeof(NotRecognizedAnnotation) && !boundAnnotationsSelections.Contains(pta.QualifiedSelection));
        }

        private IEnumerable<IParseTreeAnnotation> AttributeAnnotationsOnDeclarationsNotAllowingAttributes(IEnumerable<Declaration> userDeclarations)
        {
            return userDeclarations
                .Where(declaration => declaration.AttributesPassContext == null
                                      && !declaration.DeclarationType.HasFlag(DeclarationType.Module))
                .SelectMany(declaration => declaration.Annotations)
                .Where(pta => pta.Annotation is IAttributeAnnotation);
        }
    }
}