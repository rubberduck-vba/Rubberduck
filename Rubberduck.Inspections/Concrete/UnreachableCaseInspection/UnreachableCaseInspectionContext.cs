using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public abstract class UnreachableCaseInspectionContext
    {
        protected readonly ParserRuleContext _context;
        protected readonly IUCIValueResults _inspValues;
        protected readonly IUCIRangeClauseFilterFactory _factoryRangeClauseFilter;
        protected readonly IUCIValueFactory _factoryValue;

        public UnreachableCaseInspectionContext(ParserRuleContext context, IUCIValueResults inspValues, IUCIRangeClauseFilterFactory factory, IUCIValueFactory valueFactory)
        {
            _context = context;
            _inspValues = inspValues;
            _factoryRangeClauseFilter = factory;
            _factoryValue = valueFactory;
        }

        protected abstract bool IsResultContext<TContext>(TContext context) where TContext : ParserRuleContext;

        public TContext GetChild<TContext>() where TContext : ParserRuleContext
        {
            return Context.GetChild<TContext>();
        }

        protected ParserRuleContext Context => _context;

        protected static bool TryDetermineEvaluationTypeFromTypes(IEnumerable<string> typeNames, out string typeName)
        {
            typeName = string.Empty;
            var typeList = typeNames.ToList();

            //If everything is declared as a Variant, we do not attempt to inspect the selectStatement
            if (typeList.All(tn => new string[] { Tokens.Variant }.Contains(tn)))
            {
                return false;
            }
            typeList.All(tn => new string[] { typeList.First() }.Contains(tn));
            //If all match, the typeName is easy...This is the only way to return "String" or "Currency".
            if (typeList.All(tn => new string[] { typeList.First() }.Contains(tn)))
            {
                typeName = typeList.First();
                return true;
            }
            //Integer numbers will be evaluated using Long
            if (typeList.All(tn => new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte }.Contains(tn)))
            {
                typeName = Tokens.Long;
                return true;
            }

            //Mix of Integertypes and rational number types will be evaluated using Double
            if (typeList.All(tn => new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte, Tokens.Single, Tokens.Double }.Contains(tn)))
            {
                typeName = Tokens.Double;
                return true;
            }

            return false;
        }
    }
}
