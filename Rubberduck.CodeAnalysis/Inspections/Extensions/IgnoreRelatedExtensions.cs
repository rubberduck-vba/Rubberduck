using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;
using System.Linq;

namespace Rubberduck.Inspections.Inspections.Extensions
{
    static class IgnoreRelatedExtensions
    {
        public static bool IsIgnoringInspectionResultFor(this IdentifierReference reference, string inspectionName)
        {
            return reference.ParentScoping.HasModuleIgnoreFor(inspectionName) ||
                reference.Annotations.OfType<IgnoreAnnotation>().Any(ignore => ignore.IsIgnored(inspectionName));
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
        private static bool IsIgnoringInspectionResultFor(this QualifiedModuleName module, int line, DeclarationFinder declarationFinder, string inspectionName)
        {
            var lineScopedAnnotations = declarationFinder.FindAnnotations(module, line).OfType<IgnoreAnnotation>();
            var moduleDeclaration = declarationFinder.Members(module).First(decl => decl.DeclarationType.HasFlag(DeclarationType.Module));

            var isLineIgnored = lineScopedAnnotations.Any(annotation => annotation.IsIgnored(inspectionName));
            var isModuleIgnored = moduleDeclaration.HasModuleIgnoreFor(inspectionName);

            return isLineIgnored || isModuleIgnored;
        }
        
        private static bool HasModuleIgnoreFor(this Declaration declaration, string inspectionName)
        {
            return Declaration.GetModuleParent(declaration)?.Annotations
                .OfType<IgnoreModuleAnnotation>()
                .Any(ignoreModule => ignoreModule.IsIgnored(inspectionName)) ?? false;
        }

        private static bool HasIgnoreFor(this Declaration declaration, string inspectionName)
        {
            return declaration?.Annotations
                .OfType<IgnoreAnnotation>()
                .Any(ignore => ignore.IsIgnored(inspectionName)) ?? false;
        }
    }
}
