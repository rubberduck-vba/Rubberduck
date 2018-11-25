using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public class CodeMetricsAnalyst : ICodeMetricsAnalyst
    {

        private readonly CodeMetric[] metrics;

        public CodeMetricsAnalyst(IEnumerable<CodeMetric> supportedMetrics)
        {
            metrics = supportedMetrics.ToArray();
        }

        public IEnumerable<ICodeMetricResult> GetMetrics(RubberduckParserState state)
        {
            if (state == null || !state.AllUserDeclarations.Any())
            {
                // can not explicitly return Enumerable.Empty, this is equivalent
                yield break;
            }

            var trees = state.ParseTrees;

            foreach (var result in trees.SelectMany(moduleTree => TraverseModuleTree(moduleTree.Value, state.DeclarationFinder, moduleTree.Key)))
            {
                yield return result;
            }
        }


        public IEnumerable<ICodeMetricResult> ModuleResults(QualifiedModuleName moduleName, RubberduckParserState state)
        {
            return TraverseModuleTree(state.GetParseTree(moduleName), state.DeclarationFinder, moduleName);
        }

        private IEnumerable<ICodeMetricResult> TraverseModuleTree(IParseTree parseTree, DeclarationFinder declarationFinder, QualifiedModuleName moduleName)
        {
            var listeners = metrics.Select(metric => metric.TreeListener).ToList();
            foreach (var l in listeners)
            {
                l.Reset();
                l.InjectContext((declarationFinder, moduleName));
            }
            var combinedMetricListener = new CombinedParseTreeListener(listeners);

            ParseTreeWalker.Default.Walk(combinedMetricListener, parseTree);
            return listeners.SelectMany(l => l.Results());
        }
    }
}
