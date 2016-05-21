using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Preprocessing;
using System.Linq;

namespace Rubberduck.Parsing.Binding
{
    public static class ExpressionName
    {
        public static string GetName(VBAParser.UnrestrictedIdentifierContext context)
        {
            if (context.identifier() != null)
            {
                return GetName(context.identifier());
            }
            else
            {
                return context.GetText();
            }
        }

        public static string GetName(VBAParser.IdentifierContext context)
        {
            string name;
            var value = context.identifierValue();
            if (value.foreignName() != null)
            {
                if (value.foreignName().foreignIdentifier() != null)
                {
                    // Foreign identifiers can be nested, since the meaning of the content can differ depending on the host application,
                    // we simply everything that's inside the brackets as the identifier.
                    name = string.Join("", value.foreignName().foreignIdentifier().Select(id => id.GetText()));
                }
                else
                {
                    // Foreign identifiers can be empty, e.g. "[]".
                    name = string.Empty;
                }
            }
            else
            {
                name = value.GetText();
            }
            return name;
        }

        public static string GetName(VBAConditionalCompilationParser.NameContext context)
        {
            string name;
            var value = context.nameValue();
            if (value.foreignName() != null)
            {
                if (value.foreignName().foreignIdentifier() != null)
                {
                    // Foreign identifiers can be nested, since the meaning of the content can differ depending on the host application,
                    // we simply everything that's inside the brackets as the identifier.
                    name = string.Join("", value.foreignName().foreignIdentifier().Select(id => id.GetText()));
                }
                else
                {
                    // Foreign identifiers can be empty, e.g. "[]".
                    name = string.Empty;
                }
            }
            else
            {
                name = value.GetText();
            }
            return name;
        }
    }
}
