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
        public const string MissingInspectionSummaryElement = "MissingInspectionSummaryElement";
        private static readonly DiagnosticDescriptor MissingSummaryElementRule = new DiagnosticDescriptor(
            MissingInspectionSummaryElement,
            new LocalizableResourceString(nameof(Resources.MissingInspectionSummaryElement), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingInspectionSummaryElementMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingInspectionSummaryElementDescription), Resources.ResourceManager, typeof(Resources))
            );

        public const string MissingInspectionWhyElement = "MissingInspectionWhyElement";
        private static readonly DiagnosticDescriptor MissingWhyElementRule = new DiagnosticDescriptor(
            MissingInspectionWhyElement,
            new LocalizableResourceString(nameof(Resources.MissingInspectionWhyElement), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingInspectionWhyElementMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingInspectionWhyElementDescription), Resources.ResourceManager, typeof(Resources))
            );

        public const string MissingReferenceTag = "MissingReferenceElement";
        private static readonly DiagnosticDescriptor MissingReferenceElementRule = new DiagnosticDescriptor(
            MissingReferenceTag,
            new LocalizableResourceString(nameof(Resources.MissingInspectionReferenceElement), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingInspectionReferenceElementMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingInspectionReferenceElementDescription), Resources.ResourceManager, typeof(Resources))
        );

        public const string MissingRequiredLibraryAttribute = "MissingRequiredLibraryAttribute";
        private static readonly DiagnosticDescriptor MissingRequiredLibAttributeRule = new DiagnosticDescriptor(
            MissingRequiredLibraryAttribute,
            new LocalizableResourceString(nameof(Resources.MissingRequiredLibAttribute), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingRequiredLibAttributeMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingRequiredLibAttributeDescription), Resources.ResourceManager, typeof(Resources))
        );

        public const string MissingExampleTag = "MissingExampleElement";
        private static readonly DiagnosticDescriptor MissingExampleElementRule = new DiagnosticDescriptor(
            MissingExampleTag,
            new LocalizableResourceString(nameof(Resources.MissingExampleElement), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingExampleElementMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Warning,
            true,
            new LocalizableResourceString(nameof(Resources.MissingExampleTagDescription), Resources.ResourceManager, typeof(Resources))
        );

        public const string MissingModuleElement = "MissingModuleElement";
        private static readonly DiagnosticDescriptor MissingModuleElementRule = new DiagnosticDescriptor(
            MissingModuleElement,
            new LocalizableResourceString(nameof(Resources.MissingModuleElement), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingModuleElementMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Info,
            true,
            new LocalizableResourceString(nameof(Resources.MissingModuleElementDescription), Resources.ResourceManager, typeof(Resources))
            );

        public const string MissingNameAttribute = "MissingNameAttribute";
        private static readonly DiagnosticDescriptor MissingNameAttributeRule = new DiagnosticDescriptor(
            MissingNameAttribute,
            new LocalizableResourceString(nameof(Resources.MissingNameAttribute), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingNameAttributeMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingNameAttributeDescription), Resources.ResourceManager, typeof(Resources))
            );

        public const string MissingHasResultAttribute = "MissingHasResultAttribute";
        private static readonly DiagnosticDescriptor MissingHasResultAttributeRule = new DiagnosticDescriptor(
            MissingHasResultAttribute,
            new LocalizableResourceString(nameof(Resources.MissingHasResultAttribute), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingHasResultAttributeMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingHasResultAttributeDescription), Resources.ResourceManager, typeof(Resources))
            );

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
            ImmutableArray.Create(
                MissingSummaryElementRule, 
                MissingWhyElementRule, 
                MissingReferenceElementRule, 
                MissingRequiredLibAttributeRule,
                MissingHasResultAttributeRule,
                MissingModuleElementRule,
                MissingNameAttributeRule);

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

            var xmlTrivia = namedTypeSymbol.GetDocumentationCommentXml();
            var xml = XDocument.Parse(xmlTrivia);

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
                && string.Join(".", namedTypeSymbol.ContainingNamespace.ConstituentNamespaces).StartsWith("Rubberduck.CodeAnalysis.Inspections.Concrete")
                && namedTypeSymbol.AllInterfaces.Any(i => i.Name == "IInspection");
        }

        private static void CheckSummaryTag(SymbolAnalysisContext context, INamedTypeSymbol symbol, XDocument xml)
        {
            if (!xml.Descendants("summary").Any())
            {
                var diagnostic = Diagnostic.Create(MissingSummaryElementRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
            }
        }

        private static void CheckWhyTag(SymbolAnalysisContext context, INamedTypeSymbol symbol, XDocument xml)
        {
            if (!xml.Descendants("why").Any())
            {
                var diagnostic = Diagnostic.Create(MissingWhyElementRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
            }
        }

        private static void CheckReferenceTag(SymbolAnalysisContext context, INamedTypeSymbol symbol, XDocument xml, IEnumerable<AttributeData> requiredLibAttributes)
        {
            foreach (var referenceElement in xml.Descendants("reference"))
            {
                if (!referenceElement.Attributes().Any(a => a.Name.Equals("name")))
                {
                    var diagnostic = Diagnostic.Create(MissingNameAttributeRule, symbol.Locations[0], symbol.Name);
                    context.ReportDiagnostic(diagnostic);
                }
            }

            var xmlRefLibs = xml.Descendants("reference").Select(e => e.Attribute("name")?.Value).ToList();
            foreach (var attribute in requiredLibAttributes)
            {
                var requiredLib = attribute.ConstructorArguments[0].Value.ToString();
                if (xmlRefLibs.All(lib => lib != requiredLib))
                {
                    var diagnostic = Diagnostic.Create(MissingReferenceElementRule, symbol.Locations[0], symbol.Name, requiredLib);
                    context.ReportDiagnostic(diagnostic);
                }
            }
        }

        private static void CheckRequiredLibAttribute(SymbolAnalysisContext context, INamedTypeSymbol symbol, XDocument xml, IEnumerable<AttributeData> requiredLibAttributes)
        {
            var requiredLibs = requiredLibAttributes.Select(a => a.ConstructorArguments[0].Value.ToString()).ToList();
            foreach (var element in xml.Descendants("reference"))
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
            var examples = xml.Descendants("example");
            if (!examples.Any())
            {
                var diagnostic = Diagnostic.Create(MissingExampleElementRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
                return;
            }

            foreach (var example in examples)
            {
                if (!example.Attributes("hasresult").Any())
                {
                    var diagnostic = Diagnostic.Create(MissingHasResultAttributeRule, symbol.Locations[0], symbol.Name);
                    context.ReportDiagnostic(diagnostic);
                    return;
                }

                if (!example.Elements("module").Any())
                {
                    var diagnostic = Diagnostic.Create(MissingModuleElementRule, symbol.Locations[0], symbol.Name);
                    context.ReportDiagnostic(diagnostic);
                    return;
                }
            }
        }
    }
}
