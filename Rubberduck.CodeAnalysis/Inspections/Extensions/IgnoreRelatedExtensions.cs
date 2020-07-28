using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Extensions
{
    internal static class IgnoreRelatedExtensions
    {
        public static bool IsIgnoringInspectionResultFor(this IdentifierReference reference, string inspectionName)
        {
            return reference.ParentScoping.HasModuleIgnoreFor(inspectionName) ||
                reference.Annotations.Any(ignore => ignore.Annotation is IgnoreAnnotation && ignore.AnnotationArguments.Contains(inspectionName));
        }

        public static bool IsIgnoringInspectionResultFor(this Declaration declaration, string inspectionName)
        {
            var isIgnoredAtModuleLevel = declaration?.HasModuleIgnoreFor(inspectionName) ?? false;
            if (declaration.DeclarationType == DeclarationType.Parameter)
            {
                return isIgnoredAtModuleLevel || declaration.ParentDeclaration.HasIgnoreFor(inspectionName);
            }
            return isIgnoredAtModuleLevel || declaration.HasIgnoreFor(inspectionName);
        }

        public static bool IsIgnoringInspectionResultFor(this QualifiedContext parserContext, DeclarationFinder declarationFinder, string inspectionName)
        {
            return parserContext.ModuleName.IsIgnoringInspectionResultFor(parserContext.Context.Start.Line, declarationFinder, inspectionName);
        }

        public static bool IsIgnoringInspectionResult(this IInspectionResult result, DeclarationFinder declarationFinder)
        {
            switch (result)
            {
                case DeclarationInspectionResult declarationResult:
                    return declarationResult.Target.IsIgnoringInspectionResultFor(declarationResult.Inspection.AnnotationName);
                case IdentifierReferenceInspectionResult identifierReferenceResult:
                    return identifierReferenceResult.Reference.IsIgnoringInspectionResultFor(identifierReferenceResult.Inspection.AnnotationName);
                case QualifiedContextInspectionResult qualifiedContextResult:
                    return qualifiedContextResult.QualifiedName.IsIgnoringInspectionResultFor(qualifiedContextResult.Context.Start.Line, declarationFinder, qualifiedContextResult.Inspection.AnnotationName);
                default:
                    return false;
            }
        }

        private static bool IsIgnoringInspectionResultFor(this QualifiedModuleName module, int line, DeclarationFinder declarationFinder, string inspectionName)
        {
            var lineScopedAnnotations = declarationFinder.FindAnnotations<IgnoreAnnotation>(module, line);
            var moduleDeclaration = declarationFinder.Members(module).First(decl => decl.DeclarationType.HasFlag(DeclarationType.Module));

            var isLineIgnored = lineScopedAnnotations.Any(annotation => annotation.AnnotationArguments.Contains(inspectionName));
            var isModuleIgnored = moduleDeclaration.HasModuleIgnoreFor(inspectionName);

            return isLineIgnored || isModuleIgnored;
        }
        
        private static bool HasModuleIgnoreFor(this Declaration declaration, string inspectionName)
        {
            return Declaration.GetModuleParent(declaration)?.Annotations
                .Where(pta => pta.Annotation is IgnoreModuleAnnotation)
                .Any(ignoreModule => !ignoreModule.AnnotationArguments.Any() || ignoreModule.AnnotationArguments.Contains(inspectionName)) ?? false;
        }

        private static bool HasIgnoreFor(this Declaration declaration, string inspectionName)
        {
            return declaration?.Annotations
                .Where(pta => pta.Annotation is IgnoreAnnotation)
                .Any(ignore => ignore.AnnotationArguments.Contains(inspectionName)) ?? false;
        }
    }
}
