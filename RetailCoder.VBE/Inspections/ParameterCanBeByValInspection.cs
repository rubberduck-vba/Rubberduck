using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
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
        }

        public override string Meta { get { return InspectionsUI.ParameterCanBeByValInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ParameterCanBeByValInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var declarations = UserDeclarations.ToList();
            var issues = new List<ParameterCanBeByValInspectionResult>();

            var interfaceDeclarationMembers = declarations.FindInterfaceMembers().ToList();
            var interfaceScopes = declarations.FindInterfaceImplementationMembers().Concat(interfaceDeclarationMembers).Select(s => s.Scope);

            issues.AddRange(GetResults(declarations, interfaceDeclarationMembers));

            var eventMembers = declarations.Where(item => !item.IsBuiltIn && item.DeclarationType == DeclarationType.Event).ToList();
            var formEventHandlerScopes = State.FindFormEventHandlers().Select(handler => handler.Scope);
            var eventHandlerScopes = State.DeclarationFinder.FindEventHandlers().Concat(declarations.FindUserEventHandlers()).Select(e => e.Scope);
            var eventScopes = eventMembers.Select(s => s.Scope)
                .Concat(formEventHandlerScopes)
                .Concat(eventHandlerScopes);

            issues.AddRange(GetResults(declarations, eventMembers));

            var declareScopes = declarations.Where(item =>
                    item.DeclarationType == DeclarationType.LibraryFunction
                    || item.DeclarationType == DeclarationType.LibraryProcedure)
                .Select(e => e.Scope);
            
            issues.AddRange(declarations.Where(declaration =>
                !declaration.IsArray
                && (declaration.AsTypeDeclaration == null || declaration.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType)
                && !declareScopes.Contains(declaration.ParentScope)
                && !eventScopes.Contains(declaration.ParentScope)
                && !interfaceScopes.Contains(declaration.ParentScope)
                && declaration.DeclarationType == DeclarationType.Parameter
                && ((VBAParser.ArgContext)declaration.Context).BYVAL() == null
                && !IsUsedAsByRefParam(declarations, declaration)
                && !declaration.References.Any(reference => reference.IsAssignment))
                .Select(issue => new ParameterCanBeByValInspectionResult(this, State, issue, issue.Context, issue.QualifiedName)));

            return issues;
        }

        private IEnumerable<ParameterCanBeByValInspectionResult> GetResults(List<Declaration> declarations, List<Declaration> declarationMembers)
        {
            foreach (var declaration in declarationMembers)
            {
                var declarationParameters =
                    declarations.Where(d => d.DeclarationType == DeclarationType.Parameter &&
                                                      Equals(d.ParentDeclaration, declaration))
                                .OrderBy(o => o.Selection.StartLine)
                                .ThenBy(t => t.Selection.StartColumn)
                                .ToList();

                if (!declarationParameters.Any()) { continue; }
                var parametersAreByRef = declarationParameters.Select(s => true).ToList();

                var members = declarationMembers.Any(a => a.DeclarationType == DeclarationType.Event)
                    ? declarations.FindHandlersForEvent(declaration).Select(s => s.Item2).ToList()
                    : declarations.FindInterfaceImplementationMembers(declaration).ToList();

                foreach (var member in members)
                {
                    var parameters =
                        declarations.Where(d => d.DeclarationType == DeclarationType.Parameter &&
                                                          Equals(d.ParentDeclaration, member))
                                    .OrderBy(o => o.Selection.StartLine)
                                    .ThenBy(t => t.Selection.StartColumn)
                                    .ToList();

                    for (var i = 0; i < parameters.Count; i++)
                    {
                        parametersAreByRef[i] = parametersAreByRef[i] &&
                                                !IsUsedAsByRefParam(declarations, parameters[i]) &&
                                                ((VBAParser.ArgContext) parameters[i].Context).BYVAL() == null &&
                                                !parameters[i].References.Any(reference => reference.IsAssignment);
                    }
                }

                for (var i = 0; i < declarationParameters.Count; i++)
                {
                    if (parametersAreByRef[i])
                    {
                        yield return new ParameterCanBeByValInspectionResult(this, State, declarationParameters[i],
                            declarationParameters[i].Context, declarationParameters[i].QualifiedName);
                    }
                }
            }
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

                foreach (var declaration in calledProcedureArgs)
                {
                    if (((VBAParser.ArgContext)declaration.Context).BYVAL() != null)
                    {
                        continue;
                    }

                    if (declaration.References.Any(reference => reference.IsAssignment))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
