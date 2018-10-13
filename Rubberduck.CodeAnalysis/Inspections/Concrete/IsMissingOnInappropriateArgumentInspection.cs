using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using NLog;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    public class IsMissingOnInappropriateArgumentInspection : InspectionBase
    {
        private readonly ILogger _logger = LogManager.GetCurrentClassLogger();

        public IsMissingOnInappropriateArgumentInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var isMissing = BuiltInDeclarations.SingleOrDefault(decl => decl.QualifiedName.ToString().Equals("VBE7.DLL;VBA.Information.IsMissing"));

            if (isMissing == null)
            {
                _logger.Trace("VBA.Information.IsMissing was not found in IsMissingOnInappropriateArgumentInspection.");
                return Enumerable.Empty<IInspectionResult>();
            }

            var results = new List<IInspectionResult>();

            foreach (var reference in isMissing.References.Where(candidate => !IsIgnoringInspectionResultFor(candidate, AnnotationName)))
            {
                // First case is for unqualified use: IsMissing(foo)
                // Second case if for use as a member access: VBA.IsMissing(foo)
                var argument = ((ParserRuleContext)reference.Context.Parent).GetDescendent<VBAParser.ArgumentExpressionContext>() ??
                               ((ParserRuleContext)reference.Context.Parent.Parent).GetDescendent<VBAParser.ArgumentExpressionContext>();
                var name = argument.GetDescendent<VBAParser.SimpleNameExprContext>();
                if (name.Parent.Parent != argument)
                {
                    continue;
                }

                var procedure = reference.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();
                var parameter = UserDeclarations.Where(decl => decl is ModuleBodyElementDeclaration)
                    .Cast<ModuleBodyElementDeclaration>()
                    .FirstOrDefault(decl => decl.Context.Parent == procedure)?
                    .Parameters.FirstOrDefault(param => param.IdentifierName.Equals(name.GetText()));

                if (parameter == null || parameter.IsOptional && parameter.AsTypeName.Equals(Tokens.Variant) && string.IsNullOrEmpty(parameter.DefaultValue))
                {
                    continue;                   
                }

                results.Add(new IdentifierReferenceInspectionResult(this, InspectionResults.IsMissingOnInappropriateArgumentInspection, State, reference));
            }

            return results;
        }
    }
}
