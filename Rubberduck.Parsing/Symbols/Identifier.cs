using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Preprocessing;
using Rubberduck.VBEditor;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public static class Identifier
    {
        public static string GetName(VBAParser.FunctionNameContext context)
        {
            return GetName(context.identifier());
        }

        public static string GetName(VBAParser.SubroutineNameContext context)
        {
            return GetName(context.identifier());
        }

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
            return GetName(GetIdentifierValueContext(context));
        }

        public static string GetName(VBAParser.UntypedIdentifierContext context)
        {
            return GetName(GetIdentifierValueContext(context));
        }

        public static string GetName(VBAParser.IdentifierValueContext value)
        {
            string name;
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

        public static Selection GetNameSelection(VBAParser.UnrestrictedIdentifierContext context)
        {
            if (context.identifier() != null)
            {
                return GetNameSelection(context.identifier());
            }
            else
            {
                return context.GetSelection();
            }
        }

        public static Selection GetNameSelection(VBAParser.IdentifierContext context)
        {
            return GetIdentifierValueContext(context).GetSelection();
        }

        public static Selection GetNameSelection(VBAParser.UntypedIdentifierContext context)
        {
            return GetIdentifierValueContext(context).GetSelection();
        }

        public static VBAParser.IdentifierValueContext GetIdentifierValueContext(VBAParser.IdentifierContext context)
        {
            if (context.untypedIdentifier() != null)
            {
                return GetIdentifierValueContext(context.untypedIdentifier());
            }
            else
            {
                return context.typedIdentifier().identifierValue();
            }
        }

        public static VBAParser.IdentifierValueContext GetIdentifierValueContext(VBAParser.UntypedIdentifierContext context)
        {
            return context.identifierValue();
        }

        public static string GetTypeHintValue(VBAParser.IdentifierContext identifier)
        {
            var typeHintContext = GetTypeHintContext(identifier);
            if (typeHintContext != null)
            {
                return typeHintContext.GetText();
            }
            return null;
        }

        public static string GetTypeHintValue(VBAParser.UnrestrictedIdentifierContext identifier)
        {
            if (identifier.identifier() != null)
            {
                return GetTypeHintValue(identifier.identifier());
            }
            else
            {
                return null;
            }
        }

        public static VBAParser.TypeHintContext GetTypeHintContext(VBAParser.IdentifierContext identifier)
        {
            if (identifier.untypedIdentifier() != null)
            {
                return null;
            }
            else
            {
                return identifier.typedIdentifier().typeHint();
            }
        }
    }
}
