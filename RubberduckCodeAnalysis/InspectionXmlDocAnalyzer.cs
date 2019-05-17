using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;

namespace RubberduckCodeAnalysis
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class InspectionXmlDocAnalyzer : DiagnosticAnalyzer
    {
        private const string MissingInspectionSummaryTag = "MissingInspectionSummaryTag";
        private static readonly DiagnosticDescriptor MissingSummaryTagRule = new DiagnosticDescriptor(
            MissingInspectionSummaryTag,
            new LocalizableResourceString(nameof(Resources.MissingInspectionSummaryTag), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingInspectionSummaryTagMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingInspectionSummaryTagDescription), Resources.ResourceManager, typeof(Resources))
            );

        private const string MissingInspectionWhyTag = "MissingInspectionWhyTag";
        private static readonly DiagnosticDescriptor MissingWhyTagRule = new DiagnosticDescriptor(
            MissingInspectionWhyTag,
            new LocalizableResourceString(nameof(Resources.MissingInspectionWhyTag), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingInspectionWhyTagMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingInspectionWhyTagDescription), Resources.ResourceManager, typeof(Resources))
            );

        private const string MissingReferenceTag = "MissingReferenceTag";
        private static readonly DiagnosticDescriptor MissingReferenceTagRule = new DiagnosticDescriptor(
            MissingReferenceTag,
            new LocalizableResourceString(nameof(Resources.MissingInspectionReferenceTag), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingInspectionReferenceTagMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingInspectionReferenceTagDescription), Resources.ResourceManager, typeof(Resources))
        );

        private const string MissingRequiredLibraryAttribute = "MissingRequiredLibraryAttribute";
        private static readonly DiagnosticDescriptor MissingRequiredLibAttributeRule = new DiagnosticDescriptor(
            MissingRequiredLibraryAttribute,
            new LocalizableResourceString(nameof(Resources.MissingRequiredLibAttribute), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingRequiredLibAttributeMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingRequiredLibAttributeDescription), Resources.ResourceManager, typeof(Resources))
        );

        private const string MissingExampleTag = "MissingExampleTag";
        private static readonly DiagnosticDescriptor MissingExampleTagRule = new DiagnosticDescriptor(
            MissingExampleTag,
            new LocalizableResourceString(nameof(Resources.MissingExampleTag), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingExampleTagMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Warning,
            true,
            new LocalizableResourceString(nameof(Resources.MissingExampleTagDescription), Resources.ResourceManager, typeof(Resources))
        );

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
            ImmutableArray.Create(MissingSummaryTagRule, MissingWhyTagRule, MissingReferenceTagRule, MissingRequiredLibAttributeRule);

        public override void Initialize(AnalysisContext context)
        {
            context.RegisterSymbolAction(AnalyzeSymbol, SymbolKind.NamedType);
        }

        private static void AnalyzeSymbol(SymbolAnalysisContext context)
        {
            var namedTypeSymbol = (INamedTypeSymbol)context.Symbol;
            if (!IsInspectionClass(namedTypeSymbol))
            {
                return;
            }

            var xml = XDocument.Load(namedTypeSymbol.GetDocumentationCommentXml());

            CheckSummaryTag(context, namedTypeSymbol, xml);
            CheckWhyTag(context, namedTypeSymbol, xml);
            CheckExampleTag(context, namedTypeSymbol, xml);

            var requiredLibraryAttributes = namedTypeSymbol.GetAttributes().Where(a => a.AttributeClass.Name == "RequiredLibraryAttribute").ToList();
            CheckReferenceTag(context, namedTypeSymbol, xml, requiredLibraryAttributes);
            CheckRequiredLibAttribute(context, namedTypeSymbol, xml, requiredLibraryAttributes);
        }

        private static bool IsInspectionClass(INamedTypeSymbol namedTypeSymbol)
        {
            return namedTypeSymbol.TypeKind == TypeKind.Class && !namedTypeSymbol.IsAbstract
                && namedTypeSymbol.ContainingNamespace.Name.StartsWith("Rubberduck.CodeAnalysis.Inspections.Concrete")
                && namedTypeSymbol.AllInterfaces.Any(i => i.Name == "IInspection");
        }

        private static void CheckSummaryTag(SymbolAnalysisContext context, INamedTypeSymbol symbol, XDocument xml)
        {
            if (xml.Element("summary") == null)
            {
                var diagnostic = Diagnostic.Create(MissingSummaryTagRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
            }
        }

        private static void CheckWhyTag(SymbolAnalysisContext context, INamedTypeSymbol symbol, XDocument xml)
        {
            if (xml.Element("why") == null)
            {
                var diagnostic = Diagnostic.Create(MissingWhyTagRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
            }
        }

        private static void CheckReferenceTag(SymbolAnalysisContext context, INamedTypeSymbol symbol, XDocument xml, IEnumerable<AttributeData> requiredLibAttributes)
        {
            var xmlRefLibs = xml.Elements("reference").Select(e => e.Attribute("name")?.Value).ToList();
            foreach (var attribute in requiredLibAttributes)
            {
                var requiredLib = attribute.ConstructorArguments[0].Value.ToString();
                if (xmlRefLibs.All(lib => lib != requiredLib))
                {
                    var diagnostic = Diagnostic.Create(MissingReferenceTagRule, symbol.Locations[0], symbol.Name, requiredLib);
                    context.ReportDiagnostic(diagnostic);
                }
            }
        }

        private static void CheckRequiredLibAttribute(SymbolAnalysisContext context, INamedTypeSymbol symbol, XDocument xml, IEnumerable<AttributeData> requiredLibAttributes)
        {
            var requiredLibs = requiredLibAttributes.Select(a => a.ConstructorArguments[0].Value.ToString()).ToList();
            foreach (var element in xml.Elements("reference"))
            {
                var xmlRefLib = element.Attribute("name")?.Value;
                if (xmlRefLib == null || requiredLibs.All(lib => lib != xmlRefLib))
                {
                    var diagnostic = Diagnostic.Create(MissingRequiredLibAttributeRule, symbol.Locations[0], symbol.Name, xmlRefLib);
                    context.ReportDiagnostic(diagnostic);
                }
            }
        }

        private static void CheckExampleTag(SymbolAnalysisContext context, INamedTypeSymbol symbol, XDocument xml)
        {
            if (!xml.Elements("example").Any())
            {
                var diagnostic = Diagnostic.Create(MissingExampleTagRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
            }
        }
    }
}
