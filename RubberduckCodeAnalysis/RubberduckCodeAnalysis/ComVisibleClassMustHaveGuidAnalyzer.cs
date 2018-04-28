using System.Collections.Immutable;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;

namespace RubberduckCodeAnalysis
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class ComVisibleClassMustHaveGuidAnalyzer : DiagnosticAnalyzer
    {
        public const string DiagnosticId = "ComVisibleClassMustHaveGuidAnalyzer";
        
        private static readonly LocalizableString Title = new LocalizableResourceString(nameof(Resources.AnalyzerTitle), Resources.ResourceManager, typeof(Resources));
        private static readonly LocalizableString MessageFormat = new LocalizableResourceString(nameof(Resources.AnalyzerMessageFormat), Resources.ResourceManager, typeof(Resources));
        private static readonly LocalizableString Description = new LocalizableResourceString(nameof(Resources.AnalyzerDescription), Resources.ResourceManager, typeof(Resources));
        private static readonly LocalizableString Category = new LocalizableResourceString(nameof(Resources.AnalyzerCategory), Resources.ResourceManager, typeof(Resources));

        private static readonly DiagnosticDescriptor Rule = new DiagnosticDescriptor(DiagnosticId, Title, MessageFormat,
            Category.ToString(), DiagnosticSeverity.Error, true, Description);

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(Rule);

        public override void Initialize(AnalysisContext context)
        {
            context.RegisterSymbolAction(AnalyzeSymbol, SymbolKind.NamedType);
        }

        private static void AnalyzeSymbol(SymbolAnalysisContext context)
        {
            var namedTypeSymbol = (INamedTypeSymbol)context.Symbol;
            var attributes = namedTypeSymbol.GetAttributes();
            
            if (attributes.Any(a => a.AttributeClass.Name == nameof(ComVisibleAttribute)))
            {
                var data = attributes.Single(a => a.AttributeClass.Name == nameof(ComVisibleAttribute));
                var rawText = data.ApplicationSyntaxReference.GetSyntax().GetText();

                if (!rawText.ToString().Contains("ComVisible(true)"))
                {
                    return;
                }
            }

            if (attributes.Any(a => a.AttributeClass.Name == nameof(GuidAttribute)))
            {
                return;
            }

            var diagnostic = Diagnostic.Create(Rule, namedTypeSymbol.Locations[0], namedTypeSymbol.Name);

            context.ReportDiagnostic(diagnostic);
        }
    }
}
