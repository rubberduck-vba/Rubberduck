using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags empty code modules.
    /// </summary>
    /// <why>
    /// An empty module does not need to exist and can be safely removed.
    /// </why>
    internal sealed class EmptyModuleInspection : DeclarationInspectionBase
    {
        private readonly EmptyModuleVisitor _emptyModuleVisitor;
        private readonly IParseTreeProvider _parseTreeProvider;

        public EmptyModuleInspection(IDeclarationFinderProvider declarationFinderProvider, IParseTreeProvider parseTreeProvider)
            : base(declarationFinderProvider, new []{DeclarationType.Module}, new []{DeclarationType.Document})
        {
            _emptyModuleVisitor = new EmptyModuleVisitor();
            _parseTreeProvider = parseTreeProvider;
        }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            var module = declaration.QualifiedModuleName;
            var tree = _parseTreeProvider.GetParseTree(module, CodeKind.CodePaneCode);

            return _emptyModuleVisitor.Visit(tree);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.EmptyModuleInspection, declaration.IdentifierName);
        }
    }

    internal sealed class EmptyModuleVisitor : VBAParserBaseVisitor<bool>
    {
        //If not specified otherwise, any context makes a module non-empty.
        protected override bool DefaultResult => false;

        protected override bool AggregateResult(bool aggregate, bool nextResult)
        {
            return aggregate && nextResult;
        }

        //We bail out whenever we already know that the module is non-empty.
        protected override bool ShouldVisitNextChild(Antlr4.Runtime.Tree.IRuleNode node, bool currentResult)
        {
            return currentResult;
        }


        public override bool VisitStartRule(VBAParser.StartRuleContext context)
        {
            return Visit(context.module());
        }

        public override bool VisitModule(VBAParser.ModuleContext context)
        {
            return context.moduleConfig() == null
                && Visit(context.moduleBody())
                && Visit(context.moduleDeclarations());
        }

        public override bool VisitModuleBody(VBAParser.ModuleBodyContext context)
        {
            return !context.moduleBodyElement().Any();
        }

        public override bool VisitModuleDeclarations(VBAParser.ModuleDeclarationsContext context)
        {
            return !context.moduleDeclarationsElement().Any()
                   || context.moduleDeclarationsElement().All(Visit);
        }

        public override bool VisitModuleDeclarationsElement(VBAParser.ModuleDeclarationsElementContext context)
        {
            return context.moduleVariableStmt() == null
                   && context.moduleConstStmt() == null
                   && context.enumerationStmt() == null
                   && context.udtDeclaration() == null
                   && context.eventStmt() == null
                   && context.implementsStmt() == null
                   && context.declareStmt() == null;
        }
    }
}
