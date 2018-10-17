using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using NLog;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Inspections.Abstract
{
    public abstract class IsMissingInspectionBase : InspectionBase
    {
        private readonly ILogger _logger = LogManager.GetCurrentClassLogger();

        protected IsMissingInspectionBase(RubberduckParserState state) 
            : base(state) { }

        private static readonly List<string> IsMissingQualifiedNames = new List<string>
        {
            "VBE7.DLL;VBA.Information.IsMissing",
            "VBA6.DLL;VBA.Information.IsMissing"
        };

        protected IReadOnlyList<Declaration> IsMissingDeclarations 
        {
            get
            {
                var isMissing = BuiltInDeclarations.Where(decl => IsMissingQualifiedNames.Contains(decl.QualifiedName.ToString())).ToList();

                if (isMissing.Count == 0)
                {
                    _logger.Trace("No 'IsMissing' Declarations were not found in IsMissingInspectionBase.");
                }

                return isMissing;
            }
        }

        protected ParameterDeclaration GetParameterForReference(IdentifierReference reference)
        {
            // First case is for unqualified use: IsMissing(foo)
            // Second case if for use as a member access: VBA.IsMissing(foo)
            var argument = ((ParserRuleContext)reference.Context.Parent).GetDescendent<VBAParser.ArgumentExpressionContext>() ??
                           ((ParserRuleContext)reference.Context.Parent.Parent).GetDescendent<VBAParser.ArgumentExpressionContext>();

            var name = argument?.GetDescendent<VBAParser.SimpleNameExprContext>();
            if (name == null || name.Parent.Parent != argument)
            {
                return null;
            }

            var procedure = reference.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();
            return UserDeclarations.Where(decl => decl is ModuleBodyElementDeclaration)
                .Cast<ModuleBodyElementDeclaration>()
                .FirstOrDefault(decl => decl.Context.Parent == procedure)?
                .Parameters.FirstOrDefault(param => param.IdentifierName.Equals(name.GetText()));
        }
    }
}
