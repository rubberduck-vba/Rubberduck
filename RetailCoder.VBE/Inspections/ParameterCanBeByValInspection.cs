using System.Collections.Generic;
using System.Linq;
using System.Windows.Threading;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ParameterCanBeByValInspection : InspectionBase
    {
        public ParameterCanBeByValInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _dispatcher = Dispatcher.CurrentDispatcher;
        }

        private readonly Dispatcher _dispatcher;

        public override string Meta { get { return InspectionsUI.ParameterCanBeByValInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ParameterCanBeByValInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        // if we don't want to suggest passing non-primitive types ByRef (i.e. object types and Variant), then we need this:
        private static readonly string[] PrimitiveTypes =
        {
            Tokens.Boolean,
            Tokens.Byte,
            Tokens.Date,
            Tokens.Decimal,
            Tokens.Double,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.LongPtr,
            Tokens.Integer,
            Tokens.Single,
            Tokens.String,
            Tokens.StrPtr
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var declarations = UserDeclarations.ToList();

            IEnumerable<Declaration> interfaceMembers = null;
            interfaceMembers = declarations.FindInterfaceMembers()
                .Concat(declarations.FindInterfaceImplementationMembers())
                .ToList();

            var formEventHandlerScopes = declarations.FindFormEventHandlers()
                .Select(handler => handler.Scope);

            var eventScopes = declarations.Where(item =>
                !item.IsBuiltIn && item.DeclarationType == DeclarationType.Event)
                .Select(e => e.Scope);

            var declareScopes = declarations.Where(item =>
                    item.DeclarationType == DeclarationType.LibraryFunction
                    || item.DeclarationType == DeclarationType.LibraryProcedure)
                .Select(e => e.Scope);

            var ignoredScopes = formEventHandlerScopes.Concat(eventScopes).Concat(declareScopes);

            var issues = declarations.Where(declaration =>
                !declaration.IsArray
                && !ignoredScopes.Contains(declaration.ParentScope)
                && declaration.DeclarationType == DeclarationType.Parameter
                && !interfaceMembers.Select(m => m.Scope).Contains(declaration.ParentScope)
                && ((VBAParser.ArgContext)declaration.Context).BYVAL() == null
                && !IsUsedAsByRefParam(declarations, declaration)
                && !declaration.References.Any(reference => reference.IsAssignment))
                .Select(issue => new ParameterCanBeByValInspectionResult(this, issue, ((dynamic)issue.Context).unrestrictedIdentifier(), issue.QualifiedName));

            return issues;
        }

        private static bool IsUsedAsByRefParam(IEnumerable<Declaration> declarations, Declaration parameter)
        {
            // find the procedure calls in the procedure of the parameter.
            // note: works harder than it needs to when procedure has more than a single procedure call...
            //       ...but caching [declarations] would be a memory leak
            var items = declarations as List<Declaration> ?? declarations.ToList();

            var procedureCalls = items.Where(item => item.DeclarationType.HasFlag(DeclarationType.Member))
                .SelectMany(member => member.References.Where(reference => reference.ParentScoping.Equals(parameter.ParentScopeDeclaration)))
                .GroupBy(call => call.Declaration)
                .ToList(); // only check a procedure once. its declaration doesn't change if it's called 20 times anyway.

            foreach (var item in procedureCalls)
            {
                var calledProcedureArgs = items
                    .Where(arg => arg.DeclarationType == DeclarationType.Parameter && arg.ParentScope == item.Key.Scope)
                    .OrderBy(arg => arg.Selection.StartLine)
                    .ThenBy(arg => arg.Selection.StartColumn)
                    .ToArray();

                for (var i = 0; i < calledProcedureArgs.Count(); i++)
                {
                    if (((VBAParser.ArgContext)calledProcedureArgs[i].Context).BYVAL() != null)
                    {
                        continue;
                    }

                    foreach (var reference in item)
                    {
                        if (!(reference.Context is VBAParser.ArgContext))
                        {
                            continue;
                        }
                        var context = ((dynamic)reference.Context.Parent).argsCall() as VBAParser.ArgContext;
                        if (context == null)
                        {
                            continue;
                        }
                        if (parameter.IdentifierName == context.GetText())
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }
    }
}
