using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.JunkDrawer.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags empty code modules.
    /// </summary>
    /// <why>
    /// An empty module does not need to exist and can be safely removed.
    /// </why>
    public sealed class EmptyModuleInspection : InspectionBase
    {
        private readonly EmptyModuleVisitor _emptyModuleVisitor;

        public EmptyModuleInspection(RubberduckParserState state)
            : base(state)
        {
            _emptyModuleVisitor = new EmptyModuleVisitor();
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var modulesToInspect = State.DeclarationFinder.AllModules
                .Where(qmn => qmn.ComponentType == ComponentType.ClassModule
                        || qmn.ComponentType == ComponentType.StandardModule).ToHashSet();

            var treesToInspect = State.ParseTrees.Where(kvp => modulesToInspect.Contains(kvp.Key));

            var emptyModules = treesToInspect
                .Where(kvp => _emptyModuleVisitor.Visit(kvp.Value))
                .Select(kvp => kvp.Key)
                .ToHashSet();

            var emptyModuleDeclarations = State.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                .Where(declaration => emptyModules.Contains(declaration.QualifiedName.QualifiedModuleName));

            return emptyModuleDeclarations.Select(declaration =>
                new DeclarationInspectionResult(this, string.Format(InspectionResults.EmptyModuleInspection, declaration.IdentifierName), declaration));
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
                   && context.constStmt() == null
                   && context.enumerationStmt() == null
                   && context.udtDeclaration() == null
                   && context.eventStmt() == null
                   && context.implementsStmt() == null
                   && context.declareStmt() == null;
        }
    }
}
