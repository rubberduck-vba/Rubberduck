using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class DictionaryAccessExpression : BoundExpression
    {
        public DictionaryAccessExpression(
            Declaration referencedDeclaration,
            ExpressionClassification classification,
            ParserRuleContext context,
            IBoundExpression lExpression,
            ArgumentList argumentList)
            : base(referencedDeclaration, classification, context)
        {
            LExpression = lExpression;
            ArgumentList = argumentList;
        }

        public IBoundExpression LExpression { get; }
        public ArgumentList ArgumentList { get; }

        public ParserRuleContext DefaultMemberContext
        {
            get
            {
                if (Context is VBAParser.DictionaryAccessExprContext dictionaryAccess)
                {
                    return dictionaryAccess.dictionaryAccess();
                }

                return ((VBAParser.WithDictionaryAccessExprContext) Context).dictionaryAccess();
            }
        }
    }
}