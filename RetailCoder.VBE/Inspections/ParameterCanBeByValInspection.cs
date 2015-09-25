using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspection : IInspection
    {
        public ParameterCanBeByValInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { { return "ParameterCanBeByValInspection"; } } }
        public string Description { get { return RubberduckUI.ParameterCanBeByVal_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var interfaceMembers = parseResult.Declarations.FindInterfaceMembers()
                .Concat(parseResult.Declarations.FindInterfaceImplementationMembers())
                .ToList();

            var formEventHandlerScopes = parseResult.Declarations.FindFormEventHandlers()
                .Select(handler => handler.Scope);

            var eventScopes = parseResult.Declarations.Items.Where(item => 
                !item.IsBuiltIn && item.DeclarationType == DeclarationType.Event)
                .Select(e => e.Scope);

            var declareScopes = parseResult.Declarations.Items.Where(item => 
                    item.DeclarationType == DeclarationType.LibraryFunction 
                    || item.DeclarationType == DeclarationType.LibraryProcedure)
                .Select(e => e.Scope);

            var ignoredScopes = formEventHandlerScopes.Concat(eventScopes).Concat(declareScopes);

            var issues = parseResult.Declarations.Items.Where(declaration =>
                !ignoredScopes.Contains(declaration.ParentScope)
                && declaration.DeclarationType == DeclarationType.Parameter
                && !interfaceMembers.Select(m => m.Scope).Contains(declaration.ParentScope)
                && PrimitiveTypes.Contains(declaration.AsTypeName)
                && ((VBAParser.ArgContext) declaration.Context).BYVAL() == null
                && !IsUsedAsByRefParam(parseResult.Declarations, declaration)
                && !declaration.References.Any(reference => reference.IsAssignment))
                .Select(issue => new ParameterCanBeByValInspectionResult(this, string.Format(Description, issue.IdentifierName), ((dynamic)issue.Context).ambiguousIdentifier(), issue.QualifiedName));

            return issues;
        }

        private bool IsUsedAsByRefParam(Declarations declarations, Declaration parameter)
        {
            // find the procedure calls in the procedure of the parameter.
            // note: works harder than it needs to when procedure has more than a single procedure call...
            //       ...but caching [declarations] would be a memory leak
            var procedureCalls = declarations.Items.Where(item => item.DeclarationType.HasFlag(DeclarationType.Member))
                .SelectMany(member => member.References.Where(reference => reference.ParentScope == parameter.ParentScope))
                .GroupBy(call => call.Declaration)
                .ToList(); // only check a procedure once. its declaration doesn't change if it's called 20 times anyway.

            foreach (var item in procedureCalls)
            {
                var calledProcedureArgs = declarations.Items
                    .Where(arg => arg.DeclarationType == DeclarationType.Parameter && arg.ParentScope == item.Key.Scope)
                    .OrderBy(arg => arg.Selection.StartLine)
                    .ThenBy(arg => arg.Selection.StartColumn)
                    .ToArray();

                for (var i = 0; i < calledProcedureArgs.Count(); i++)
                {
                    if (((VBAParser.ArgContext) calledProcedureArgs[i].Context).BYVAL() == null)
                    {
                        foreach (var reference in item)
                        {
                            var context = ((dynamic)reference.Context.Parent).argsCall() as VBAParser.ArgsCallContext;
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
            }

            return false;
        }
    }
}