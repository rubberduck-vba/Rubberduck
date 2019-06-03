using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;

namespace RubberduckCodeAnalysis
{

    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class ComVisibleTypeAnalyzer : DiagnosticAnalyzer
    {
        private const string MissingGuidId = "MissingGuid";
        private static readonly DiagnosticDescriptor MissingGuidRule = new DiagnosticDescriptor(
            MissingGuidId,
            new LocalizableResourceString(nameof(Resources.MissingGuidTitle), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingGuidMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.AnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(), 
            DiagnosticSeverity.Error, 
            true,
            new LocalizableResourceString(nameof(Resources.MissingGuidDescription), Resources.ResourceManager, typeof(Resources))
        );

        private const string MissingClassInterfaceId = "MissingClassInterface";
        private static readonly DiagnosticDescriptor MissingClassInterfaceRule = new DiagnosticDescriptor(
            MissingClassInterfaceId,
            new LocalizableResourceString(nameof(Resources.MissingClassInterfaceTitle), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingClassInterfaceMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.AnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingClassInterfaceDescription), Resources.ResourceManager, typeof(Resources))
        );

        private const string MissingProgIdId = "MissingProgId";
        private static readonly DiagnosticDescriptor MissingProgIdRule = new DiagnosticDescriptor(
            MissingProgIdId,
            new LocalizableResourceString(nameof(Resources.MissingProgIdTitle), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingProgIdMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.AnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingProgIdDescription), Resources.ResourceManager, typeof(Resources))
        );

        private const string MissingComDefaultInterfaceId = "MissingComDefaultInterface";
        private static readonly DiagnosticDescriptor MissingComDefaultInterfaceRule = new DiagnosticDescriptor(
            MissingComDefaultInterfaceId,
            new LocalizableResourceString(nameof(Resources.MissingComDefaultInterfaceTitle), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingComDefaultInterfaceMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.AnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingComDefaultInterfaceDescription), Resources.ResourceManager, typeof(Resources))
        );

        private const string MissingInterfaceTypeId = "MissingInterfaceType";
        private static readonly DiagnosticDescriptor MissingInterfaceTypeRule = new DiagnosticDescriptor(
            MissingInterfaceTypeId,
            new LocalizableResourceString(nameof(Resources.MissingInterfaceTypeTitle), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingInterfaceTypeMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.AnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingInterfaceTypeDescription), Resources.ResourceManager, typeof(Resources))
        );

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
            ImmutableArray.Create(MissingGuidRule, MissingClassInterfaceRule, MissingProgIdRule,
                MissingComDefaultInterfaceRule,
                MissingInterfaceTypeRule);

        public override void Initialize(AnalysisContext context)
        {
            context.RegisterSymbolAction(AnalyzeSymbol, SymbolKind.NamedType);
        }

        private static void AnalyzeSymbol(SymbolAnalysisContext context)
        {
            var namedTypeSymbol = (INamedTypeSymbol) context.Symbol;
            var attributes = namedTypeSymbol.GetAttributes();

            if (!IsComVisibleType(attributes))
            {
                return;
            }

            CheckGuidAttribute(context, namedTypeSymbol, attributes);

            if (IsClass(namedTypeSymbol))
            {
                CheckProgIdAttribute(context, namedTypeSymbol, attributes);
                CheckClassInterfaceAttribute(context, namedTypeSymbol, attributes);
                CheckComDefaultInterfaceAttribute(context, namedTypeSymbol, attributes);
            }

            if (IsInterface(namedTypeSymbol))
            {
                CheckInterfaceTypeAttribute(context, namedTypeSymbol, attributes);
            }
        }

        private static bool IsComVisibleType(IEnumerable<AttributeData> attributes)
        {
            if (!attributes.Any(a => a.AttributeClass.Name == nameof(ComVisibleAttribute)))
            {
                return false;
            }
            var data = attributes.Single(a => a.AttributeClass.Name == nameof(ComVisibleAttribute));
            var rawText = data.ApplicationSyntaxReference.GetSyntax().GetText();

            if (!rawText.ToString().Contains("ComVisible(true)"))
            {
                return false;
            }

            return true;
        }

        private static bool IsClass(INamedTypeSymbol namedTypeSymbol)
        {
            return namedTypeSymbol.TypeKind == TypeKind.Class;
        }

        private static bool IsInterface(INamedTypeSymbol namedTypeSymbol)
        {
            return namedTypeSymbol.TypeKind == TypeKind.Interface;
        }

        private static void CheckGuidAttribute(SymbolAnalysisContext context, INamedTypeSymbol namedTypeSymbol, IEnumerable<AttributeData> attributes)
        {
            if (attributes.Any(a => a.AttributeClass.Name == nameof(GuidAttribute)))
            {
                var data = attributes.Single(a => a.AttributeClass.Name == nameof(GuidAttribute));
                var rawText = data.ApplicationSyntaxReference.GetSyntax().GetText();

                if (rawText.ToString().Contains("Guid(RubberduckGuid."))
                {
                    return;
                }
            }

            var diagnostic = Diagnostic.Create(MissingGuidRule, namedTypeSymbol.Locations[0], namedTypeSymbol.Name);
            context.ReportDiagnostic(diagnostic);
        }

        private static void CheckClassInterfaceAttribute(SymbolAnalysisContext context,
            INamedTypeSymbol namedTypeSymbol, IEnumerable<AttributeData> attributes)
        {
            if (attributes.Any(a => a.AttributeClass.Name == nameof(ClassInterfaceAttribute)))
            {
                var data = attributes.Single(a => a.AttributeClass.Name == nameof(ClassInterfaceAttribute));
                var rawText = data.ApplicationSyntaxReference.GetSyntax().GetText();

                if (rawText.ToString().Contains("ClassInterface(ClassInterfaceType.None)"))
                {
                    return;
                }
            }

            var diagnostic = Diagnostic.Create(MissingClassInterfaceRule, namedTypeSymbol.Locations[0], namedTypeSymbol.Name);
            context.ReportDiagnostic(diagnostic);
        }
        
        private static void CheckProgIdAttribute(SymbolAnalysisContext context,
            INamedTypeSymbol namedTypeSymbol, IEnumerable<AttributeData> attributes)
        {
            // We can't use nameof because..... I don't know.
            if (attributes.Any(a => a.AttributeClass.Name == "ProgIdAttribute"))
            {
                var data = attributes.Single(a => a.AttributeClass.Name == "ProgIdAttribute");
                var rawText = data.ApplicationSyntaxReference.GetSyntax().GetText();

                if (rawText.ToString().Contains("ProgId(RubberduckProgId."))
                {
                    return;
                }
            }

            var diagnostic = Diagnostic.Create(MissingProgIdRule, namedTypeSymbol.Locations[0], namedTypeSymbol.Name);
            context.ReportDiagnostic(diagnostic);
        }

        private static void CheckComDefaultInterfaceAttribute(SymbolAnalysisContext context,
            INamedTypeSymbol namedTypeSymbol, IEnumerable<AttributeData> attributes)
        {
            if (attributes.Any(a => a.AttributeClass.Name == nameof(ComDefaultInterfaceAttribute)))
            {
                var data = attributes.Single(a => a.AttributeClass.Name == nameof(ComDefaultInterfaceAttribute));
                var rawText = data.ApplicationSyntaxReference.GetSyntax().GetText();

                if (rawText.ToString().Contains("ComDefaultInterface(typeof("))
                {
                    return;
                }
            }

            var diagnostic = Diagnostic.Create(MissingComDefaultInterfaceRule, namedTypeSymbol.Locations[0], namedTypeSymbol.Name);
            context.ReportDiagnostic(diagnostic);
        }

        private static void CheckInterfaceTypeAttribute(SymbolAnalysisContext context,
            INamedTypeSymbol namedTypeSymbol, IEnumerable<AttributeData> attributes)
        {
            if (attributes.Any(a => a.AttributeClass.Name == nameof(InterfaceTypeAttribute)))
            {
                var data = attributes.Single(a => a.AttributeClass.Name == nameof(InterfaceTypeAttribute));
                var rawText = data.ApplicationSyntaxReference.GetSyntax().GetText();

                if (rawText.ToString().Contains("InterfaceType(ComInterfaceType."))
                {
                    return;
                }
            }

            var diagnostic = Diagnostic.Create(MissingInterfaceTypeRule, namedTypeSymbol.Locations[0], namedTypeSymbol.Name);
            context.ReportDiagnostic(diagnostic);
        }

    }
}
