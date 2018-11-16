using System.Collections.Immutable;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

namespace RubberduckCodeAnalysis
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class ChainedWrapperAnalyzer : DiagnosticAnalyzer
    {
        private const string ChainedWrapperId = "ChainedWrapper";
        private static readonly DiagnosticDescriptor ChainedWrapperRule = new DiagnosticDescriptor(
            ChainedWrapperId,
            new LocalizableResourceString(nameof(Resources.ChainedWrapperTitle), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.ChainedWrapperMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.AnalyzerCategory), Resources.ResourceManager, typeof(Resources)).ToString(), 
            DiagnosticSeverity.Error, 
            true,
            new LocalizableResourceString(nameof(Resources.ChainedWrapperDescription), Resources.ResourceManager, typeof(Resources))
        );

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
            ImmutableArray.Create(ChainedWrapperRule);

        public override void Initialize(AnalysisContext context)
        {
            context.RegisterSyntaxNodeAction(AnalyzeSymbol, SyntaxKind.SimpleMemberAccessExpression);
        }
       
        private static void AnalyzeSymbol(SyntaxNodeAnalysisContext context)
        {
            var node = (MemberAccessExpressionSyntax)context.Node;

            if (!(node.Expression is InvocationExpressionSyntax || node.Expression is MemberAccessExpressionSyntax))
            {
                return;
            }

            var expInterfaces = context.SemanticModel.GetTypeInfo(node.Expression).Type?.AllInterfaces;

            var nameValue = node.Name.Parent.Parent is InvocationExpressionSyntax ? node.Name.Parent.Parent : node.Name;
            var nameInterfaces = context.SemanticModel.GetTypeInfo(nameValue).Type?.AllInterfaces;

            if (!expInterfaces.HasValue || !nameInterfaces.HasValue)
            {
                return;
            }

            if (expInterfaces.Value.Any(a => a.ToDisplayString() == "Rubberduck.VBEditor.SafeComWrappers.Abstract.ISafeComWrapper") &&
                nameInterfaces.Value.Any(a => a.ToDisplayString() == "Rubberduck.VBEditor.SafeComWrappers.Abstract.ISafeComWrapper"))
            {
                var targetType = context.SemanticModel.GetTypeInfo(nameValue).Type.Name;
                var containingType = context.SemanticModel.GetTypeInfo(node.Expression).Type.Name;
                var diagnostic = Diagnostic.Create(ChainedWrapperRule, node.GetLocation(), targetType, containingType, node.GetText());
                
                context.ReportDiagnostic(diagnostic);
            }
        }
    }
}
