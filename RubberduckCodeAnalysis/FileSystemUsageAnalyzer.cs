using System.Collections.Immutable;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Text;

namespace RubberduckCodeAnalysis
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public class FileSystemUsageAnalyzer : DiagnosticAnalyzer
    {
        public const string FileSystemUsageId = "FileSystemUsageAnalyzer";
        private static readonly DiagnosticDescriptor FileSystemUsageRule = new DiagnosticDescriptor(
            FileSystemUsageId,
            new LocalizableResourceString(nameof(Resources.FileSystemUsageTile), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.FileSystemUsageMessageFormat), Resources.ResourceManager, typeof(Resources)),
            new LocalizableResourceString(nameof(Resources.FileSystemUsageCategory), Resources.ResourceManager, typeof(Resources)).ToString(),
            DiagnosticSeverity.Error,
            true,
            null
            );

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(
            FileSystemUsageRule
            );

        public override void Initialize(AnalysisContext context)
        {
            context.RegisterSemanticModelAction(Analyze);
        }

        private static void Analyze(SemanticModelAnalysisContext context)
        {
            var tree = context.SemanticModel.SyntaxTree;
            TextLine line = default;
            if(tree.GetText().Lines.Any(l =>
            {
                if (l.ToString().Equals("using System.IO;"))
                {
                    line = l;
                    return true;
                }
                return false;
            }))
            {
                var diagnostic = Diagnostic.Create(FileSystemUsageRule, tree.GetLocation(line.Span), tree.FilePath);
                context.ReportDiagnostic(diagnostic);
            }
        }
    }
}
