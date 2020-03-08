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

        public const string MissingReferenceElement = "MissingReferenceElement";
        private static readonly DiagnosticDescriptor MissingReferenceElementRule = new DiagnosticDescriptor(
            MissingReferenceElement,
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

        public const string MissingExampleElement = "MissingExampleElement";
        private static readonly DiagnosticDescriptor MissingExampleElementRule = new DiagnosticDescriptor(
            MissingExampleElement,
            new LocalizableResourceString(nameof(Resources.MissingExampleElement), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingExampleElementMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Warning,
            true,
            new LocalizableResourceString(nameof(Resources.MissingExampleElementDescription), Resources.ResourceManager, typeof(Resources))
        );

        public const string MissingModuleElement = "MissingModuleElement";
        private static readonly DiagnosticDescriptor MissingModuleElementRule = new DiagnosticDescriptor(
            MissingModuleElement,
            new LocalizableResourceString(nameof(Resources.MissingModuleElement), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingModuleElementMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
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

        public const string DuplicateNameAttribute = "DuplicateNameAttribute";
        private static readonly DiagnosticDescriptor DuplicateNameAttributeRule = new DiagnosticDescriptor(
            DuplicateNameAttribute,
            new LocalizableResourceString(nameof(Resources.DuplicateNameAttribute), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.DuplicateNameAttributeMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.DuplicateNameAttributeDescription), Resources.ResourceManager, typeof(Resources))
        );

        public const string MissingTypeAttribute = "MissingTypeAttribute";
        private static readonly DiagnosticDescriptor MissingTypeAttributeRule = new DiagnosticDescriptor(
            MissingTypeAttribute,
            new LocalizableResourceString(nameof(Resources.MissingTypeAttribute), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.MissingTypeAttributeMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.MissingTypeAttributeDescription), Resources.ResourceManager, typeof(Resources))
        );

        public const string InvalidTypeAttribute = "InvalidTypeAttribute";
        private static readonly DiagnosticDescriptor InvalidTypeAttributeRule = new DiagnosticDescriptor(
            InvalidTypeAttribute,
            new LocalizableResourceString(nameof(Resources.InvalidTypeAttribute), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.InvalidTypeAttributeMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.XmlDocAnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            new LocalizableResourceString(nameof(Resources.InvalidTypeAttributeDescription), Resources.ResourceManager, typeof(Resources))
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

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(
            MissingSummaryElementRule, 
            MissingWhyElementRule, 
            MissingReferenceElementRule, 
            MissingRequiredLibAttributeRule,
            MissingHasResultAttributeRule,
            MissingNameAttributeRule,
            MissingModuleElementRule,
            MissingExampleElementRule,
            MissingTypeAttributeRule,
            InvalidTypeAttributeRule,
            DuplicateNameAttributeRule
            );

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

            var xml = XDocument.Parse(namedTypeSymbol.GetDocumentationCommentXml()).Element("member");

            CheckSummaryElement(context, namedTypeSymbol, xml);
            CheckWhyElement(context, namedTypeSymbol, xml);
            CheckExampleElement(context, namedTypeSymbol, xml);

            var requiredLibraryAttributes = namedTypeSymbol.GetAttributes().Where(a => a.AttributeClass.Name == "RequiredLibraryAttribute").ToList();
            CheckReferenceElement(context, namedTypeSymbol, xml, requiredLibraryAttributes);
            CheckRequiredLibAttribute(context, namedTypeSymbol, xml, requiredLibraryAttributes);
        }

        private static bool IsInspectionClass(INamedTypeSymbol namedTypeSymbol)
        {
            return namedTypeSymbol.TypeKind == TypeKind.Class && !namedTypeSymbol.IsAbstract
                && namedTypeSymbol.ContainingNamespace.ToString().StartsWith("Rubberduck.CodeAnalysis.Inspections.Concrete")
                && namedTypeSymbol.AllInterfaces.Any(i => i.Name == "IInspection");
        }

        private static void CheckSummaryElement(SymbolAnalysisContext context, INamedTypeSymbol symbol, XElement xml)
        {
            if (xml.Element("summary") == null)
            {
                var diagnostic = Diagnostic.Create(MissingSummaryElementRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
            }
        }

        private static void CheckWhyElement(SymbolAnalysisContext context, INamedTypeSymbol symbol, XElement xml)
        {
            if (xml.Element("why") == null)
            {
                var diagnostic = Diagnostic.Create(MissingWhyElementRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
            }
        }

        private static string CheckNameAttributeAndReturnValue(SymbolAnalysisContext context, XElement element, Location location)
        {
            var nameAttribute = element.Attributes().FirstOrDefault(a => a.Name.LocalName.Equals("name"));
            if (nameAttribute == null)
            {
                var diagnostic = Diagnostic.Create(MissingNameAttributeRule, location, element.Name.LocalName);
                context.ReportDiagnostic(diagnostic);
            }

            return nameAttribute?.Value;
        }

        private static void CheckReferenceElement(SymbolAnalysisContext context, INamedTypeSymbol symbol, XElement xml, ICollection<AttributeData> requiredLibAttributes)
        {
            if (requiredLibAttributes.Any() && !xml.Elements("reference").Any())
            {
                var diagnostic = Diagnostic.Create(MissingReferenceElementRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
            }

            var xmlRefLibs = new List<string>();
            foreach (var element in xml.Elements("reference"))
            {
                var name = CheckNameAttributeAndReturnValue(context, element, symbol.Locations[0]);
                if (name != null)
                {
                    xmlRefLibs.Add(name);
                }
            }

            var duplicateNames = xmlRefLibs
                .GroupBy(name => name)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key);
            foreach (var name in duplicateNames)
            {
                var diagnostic = Diagnostic.Create(DuplicateNameAttributeRule, symbol.Locations[0], name, "reference");
                context.ReportDiagnostic(diagnostic);
            }
            
            foreach (var attribute in requiredLibAttributes)
            {
                var requiredLib = attribute.ConstructorArguments[0].Value.ToString();
                if (xmlRefLibs.All(lib => lib != requiredLib))
                {
                    var diagnostic = Diagnostic.Create(MissingReferenceElementRule, symbol.Locations[0], symbol.Name);
                    context.ReportDiagnostic(diagnostic);
                }
            }
        }

        private static void CheckRequiredLibAttribute(SymbolAnalysisContext context, INamedTypeSymbol symbol, XElement xml, IEnumerable<AttributeData> requiredLibAttributes)
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

        private static void CheckExampleElement(SymbolAnalysisContext context, INamedTypeSymbol symbol, XElement xml)
        {
            if (!xml.Elements("example").Any())
            {
                var diagnostic = Diagnostic.Create(MissingExampleElementRule, symbol.Locations[0], symbol.Name);
                context.ReportDiagnostic(diagnostic);
                return;
            }

            var examples = xml.Elements("example");
            foreach (var example in examples)
            {
                CheckHasResultAttribute(context, example, symbol.Locations[0]);
                CheckModuleElements(context, symbol, example);
            }
        }

        private static void CheckModuleElements(SymbolAnalysisContext context, INamedTypeSymbol symbol, XElement example)
        {
            if (!example.Elements("module").Any())
            {
                var diagnostic = Diagnostic.Create(MissingModuleElementRule, symbol.Locations[0]);
                context.ReportDiagnostic(diagnostic);
            }

            var moduleNames = new List<string>();
            foreach (var module in example.Elements("module"))
            {
                var moduleName = CheckNameAttributeAndReturnValue(context, module, symbol.Locations[0]);
                if (moduleName != null)
                {
                    moduleNames.Add(moduleName);
                }

                CheckTypeAttribute(context, module, symbol.Locations[0]);
            }

            var duplicateNames = moduleNames
                .GroupBy(name => name)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key);
            foreach (var name in duplicateNames)
            {
                var diagnostic = Diagnostic.Create(DuplicateNameAttributeRule, symbol.Locations[0], name, "module");
                context.ReportDiagnostic(diagnostic);
            }
        }

        private static void CheckHasResultAttribute(SymbolAnalysisContext context, XElement element, Location location)
        {
            if (!element.Attributes().Any(a => a.Name.LocalName.Equals("hasresult", System.StringComparison.InvariantCultureIgnoreCase)))
            {
                var diagnostic = Diagnostic.Create(MissingHasResultAttributeRule, location);
                context.ReportDiagnostic(diagnostic);
            }
        }

        private static void CheckTypeAttribute(SymbolAnalysisContext context, XElement element, Location location)
        {
            var nameAttribute = element.Attributes().FirstOrDefault(a => a.Name.LocalName.Equals("type"));
            if (nameAttribute == null)
            {
                var diagnostic = Diagnostic.Create(MissingTypeAttributeRule, location, element.Name.LocalName);
                context.ReportDiagnostic(diagnostic);
            }
            else
            {
                var typeNameValue = nameAttribute.Value;
                if (!ValidTypeAttributeValues.Contains(typeNameValue))
                {
                    var diagnostic = Diagnostic.Create(InvalidTypeAttributeRule, location, typeNameValue);
                    context.ReportDiagnostic(diagnostic);
                }
            }
        }

        private static readonly List<string> ValidTypeAttributeValues = new List<string>
        {
            "Standard Module",
            "Class Module",
            "Document",
            "User Form"
        };
    }
}
