using System.Diagnostics;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.PreProcessing;

namespace Rubberduck.Parsing.Symbols
{
    public static class Identifier
    {

        public static string GetName(VBAParser.SubStmtContext context, out Interval tokenInterval)
        {
            var nameContext = context.subroutineName();
            return GetName(nameContext, out tokenInterval);
        }

        public static string GetName(VBAParser.FunctionStmtContext context, out Interval tokenInterval)
        {
            var nameContext = context.functionName();
            return GetName(nameContext, out tokenInterval);
        }

        public static string GetName(VBAParser.EventStmtContext context, out Interval tokenInterval)
        {
            var nameContext = context.identifier();
            return GetName(nameContext, out tokenInterval);
        }

        public static string GetName(VBAParser.VariableSubStmtContext context, out Interval tokenInterval)
        {
            var nameContext = context.identifier();
            return GetName(nameContext, out tokenInterval);
        }

        public static string GetName(VBAParser.PropertyGetStmtContext context, out Interval tokenInterval)
        {
            var nameContext = context.functionName();
            return GetName(nameContext, out tokenInterval);
        }

        public static string GetName(VBAParser.PropertyLetStmtContext context, out Interval tokenInterval)
        {
            var nameContext = context.subroutineName();
            return GetName(nameContext, out tokenInterval);
        }

        public static string GetName(VBAParser.PropertySetStmtContext context, out Interval tokenInterval)
        {
            var nameContext = context.subroutineName();
            return GetName(nameContext, out tokenInterval);
        }

        public static string GetName(VBAParser.ArgContext context, out Interval tokenInterval)
        {
            var nameContext = context.unrestrictedIdentifier();
            return GetName(nameContext, out tokenInterval);
        }

        public static string GetName(VBAParser.FunctionNameContext context, out Interval tokenInterval)
        {
            var nameContext = context.identifier();
            tokenInterval = Interval.Of(nameContext.Start.TokenIndex, nameContext.Stop.TokenIndex);
            return GetName(context);
        }

        public static string GetName(VBAParser.FunctionNameContext context)
        {
            return GetName(context.identifier());
        }

        public static string GetName(VBAParser.SubroutineNameContext context, out Interval tokenInterval)
        {
            var nameContext = context.identifier();
            tokenInterval = Interval.Of(nameContext.Start.TokenIndex, nameContext.Stop.TokenIndex);
            return GetName(context);
        }

        public static string GetName(VBAParser.SubroutineNameContext context)
        {
            return GetName(context.identifier());
        }

        public static string GetName(VBAParser.UnrestrictedIdentifierContext context, out Interval tokenInterval)
        {
            var nameContext = context.identifier();
            tokenInterval = Interval.Of(nameContext.Start.TokenIndex, nameContext.Stop.TokenIndex);
            return GetName(context);
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

        public static string GetName(VBAParser.IdentifierContext context, out Interval tokenInterval)
        {
            tokenInterval = Interval.Of(context.Start.TokenIndex, context.Stop.TokenIndex);
            return GetName(context);
        }

        public static string GetName(VBAParser.IdentifierContext context)
        {
            return GetName(GetIdentifierValueContext(context));
        }

        public static string GetName(VBAParser.UntypedIdentifierContext context, out Interval tokenInterval)
        {
            tokenInterval = Interval.Of(context.Start.TokenIndex, context.Stop.TokenIndex);
            return GetName(context);
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
                return context.typedIdentifier().untypedIdentifier().identifierValue();
            }
        }

        public static VBAParser.IdentifierValueContext GetIdentifierValueContext(VBAParser.UntypedIdentifierContext context)
        {
            return context.identifierValue();
        }

        public static string GetTypeHintValue(VBAParser.IdentifierContext identifier)
        {
            var typeHintContext = GetTypeHintContext(identifier);
            return typeHintContext?.GetText();
        }

        public static string GetTypeHintValue(VBAParser.IdentifierContext identifier, out IToken token)
        {
            var typeHintContext = GetTypeHintContext(identifier);
            token = typeHintContext.Start;
            Debug.Assert(typeHintContext.Stop.TokenIndex == token.TokenIndex);
            return typeHintContext.GetText();
        }

        public static string GetTypeHintValue(VBAParser.UnrestrictedIdentifierContext identifier)
        {
            return identifier.identifier() != null 
                ? GetTypeHintValue(identifier.identifier()) 
                : null;
        }

        public static string GetTypeHintValue(VBAParser.UnrestrictedIdentifierContext identifier, out IToken token)
        {
            token = null;
            return identifier.identifier() != null
                ? GetTypeHintValue(identifier.identifier(), out token)
                : null;
        }

        public static VBAParser.TypeHintContext GetTypeHintContext(VBAParser.IdentifierContext identifier)
        {
            return identifier.untypedIdentifier() != null 
                ? null 
                : identifier.typedIdentifier().typeHint();
        }
    }
}
