using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ConvertToProcedureQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(NonReturningFunctionInspection),
            typeof(FunctionReturnValueNotUsedInspection)
        };

        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        public void Fix(IInspectionResult result)
        {
            dynamic functionContext = result.Target.Context as VBAParser.FunctionStmtContext;
            dynamic propertyGetContext = result.Target.Context as VBAParser.PropertyGetStmtContext;

            var context = functionContext ?? propertyGetContext;
            if (context == null)
            {
                throw new InvalidOperationException(string.Format(InspectionsUI.InvalidContextTypeInspectionFix, result.Target.Context.GetType(), GetType()));
            }

            var functionName = result.Target.Context is VBAParser.FunctionStmtContext
                ? ((VBAParser.FunctionStmtContext)result.Target.Context).functionName()
                : ((VBAParser.PropertyGetStmtContext)result.Target.Context).functionName();

            var token = functionContext != null
                ? Tokens.Function
                : Tokens.Property + ' ' + Tokens.Get;
            var endToken = token == Tokens.Function
                ? token
                : Tokens.Property;

            string visibility = context.visibility() == null ? string.Empty : context.visibility().GetText() + ' ';
            var name = ' ' + Identifier.GetName(functionName.identifier());
            var hasTypeHint = Identifier.GetTypeHintValue(functionName.identifier()) != null;

            string args = context.argList().GetText();
            string asType = context.asTypeClause() == null ? string.Empty : ' ' + context.asTypeClause().GetText();

            var oldSignature = visibility + token + name + (hasTypeHint ? Identifier.GetTypeHintValue(functionName.identifier()) : string.Empty) + args + asType;
            var newSignature = visibility + Tokens.Sub + name + args;

            var procedure = result.Target.Context.GetText();
            var noReturnStatements = procedure;

            GetReturnStatements(result.Target).ToList().ForEach(returnStatement =>
                noReturnStatements = Regex.Replace(noReturnStatements, @"[ \t\f]*" + returnStatement + @"[ \t\f]*\r?\n?", ""));
            var newText = noReturnStatements.Replace(oldSignature, newSignature)
                .Replace(Tokens.End + ' ' + endToken, Tokens.End + ' ' + Tokens.Sub)
                .Replace(Tokens.Exit + ' ' + endToken, Tokens.Exit + ' ' + Tokens.Sub);

            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            var selection = result.Target.Context.GetSelection();

            module.DeleteLines(selection.StartLine, selection.LineCount);
            module.InsertLines(selection.StartLine, newText);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ConvertFunctionToProcedureQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => false;

        private IEnumerable<string> GetReturnStatements(Declaration declaration)
        {
            return declaration.References
                .Where(usage => IsReturnStatement(declaration, usage))
                .Select(usage => usage.Context.Parent.GetText());
        }

        private bool IsReturnStatement(Declaration declaration, IdentifierReference assignment)
        {
            return assignment.ParentScoping.Equals(declaration) && assignment.Declaration.Equals(declaration);
        }
    }
}
