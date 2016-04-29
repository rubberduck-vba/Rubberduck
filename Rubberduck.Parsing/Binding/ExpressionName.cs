namespace Rubberduck.Parsing.Binding
{
    public static class ExpressionName
    {
        public static string GetName(VBAExpressionParser.NameContext context)
        {
            string name;
            if (context.untypedName() != null)
            {
                name = context.untypedName().GetText();
            }
            else
            {
                name = context.typedName().typedNameValue().GetText();
            }
            return name;
        }

        public static string GetName(VBAExpressionParser.ReservedIdentifierNameContext context)
        {
            string name;
            if (context.reservedUntypedName() != null)
            {
                name = context.reservedUntypedName().GetText();
            }
            else
            {
                name = context.reservedTypedName().reservedIdentifier().GetText();
            }
            return name;
        }

        public static string GetName(VBAExpressionParser.UnrestrictedNameContext context)
        {
            if (context.name() != null)
            {
                return GetName(context.name());
            }
            else
            {
                return GetName(context.reservedIdentifierName());
            }
        }
    }
}
