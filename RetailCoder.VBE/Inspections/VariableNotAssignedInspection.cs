using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class VariableNotAssignedInspection : IInspection
    {
        public VariableNotAssignedInspection()
        {
            Severity = CodeInspectionSeverity.Error;
        }

        public string Name { get { return InspectionNames.VariableNotAssigned; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            var parseResults = parseResult.ToList();

            // publics & globals delared at module-scope in standard modules:
            var globals = 
                parseResults.Select(result => 
                                new 
                                { 
                                    Name = result.QualifiedName, 
                                    Globals = result.ParseTree.GetContexts<DeclarationSectionListener, ParserRuleContext>(new DeclarationSectionListener())
                                                              .OfType<VisualBasic6Parser.VariableStmtContext>()
                                                              .Where(context => 
                                                                  context.visibility() != null && 
                                                                  context.visibility().GetText() != Tokens.Private)
                                                              .SelectMany(context => context.variableListStmt().variableSubStmt()
                                                                                    .Select(variable => variable.ambiguousIdentifier()))
                                })
                             .SelectMany(module => module.Globals.Select(global => 
                                new
                                {
                                    Name = module.Name,
                                    Global = global
                                })).ToList();

            var assignedGlobals = new List<VisualBasic6Parser.AmbiguousIdentifierContext>();
            var unassignedDeclarations = new List<CodeInspectionResultBase>();

            foreach (var result in parseResults)
            {
                // module-scoped in this module:
                var declarations = result.ParseTree.GetContexts<DeclarationSectionListener, ParserRuleContext>(new DeclarationSectionListener())
                                         .OfType<VisualBasic6Parser.VariableSubStmtContext>()
                                         .Where(variable => globals.All(global => global.Global.GetText() != variable.GetText()))
                                         .ToList();
                var procedures = result.ParseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener()).ToList();

                // fetch & scope all assignments:
                var assignments = procedures.SelectMany(
                    procedure => procedure.GetContexts<VariableAssignmentListener, VisualBasic6Parser.AmbiguousIdentifierContext>(new VariableAssignmentListener())
                                         .Select(context => new
                                             {
                                                 Scope = new QualifiedMemberName(result.QualifiedName, ((dynamic)procedure).ambiguousIdentifier().GetText()),
                                                 Name = context.GetText()
                                             }));

                // fetch & scope all procedure-scoped declarations:
                var locals = procedures.SelectMany(
                    procedure => procedure.GetContexts<DeclarationListener, ParserRuleContext>(new DeclarationListener())
                                          .OfType<VisualBasic6Parser.VariableSubStmtContext>()
                                          .Select(context => new
                                             {
                                                 Context = context,
                                                 Scope = new QualifiedMemberName(result.QualifiedName, ((dynamic)procedure).ambiguousIdentifier().GetText()),
                                                 Name = context.ambiguousIdentifier().GetText()
                                             }));

                // identify unassigned module-scoped declarations:
                unassignedDeclarations.AddRange(
                    declarations.Select(d => d.ambiguousIdentifier())
                                .Where(d => globals.All(g => g.Global.GetText() != d.GetText()) 
                                        && assignments.All(a => a.Name != d.GetText()))
                                .Select(identifier => new VariableNotAssignedInspectionResult(Name, Severity, identifier, result.QualifiedName)));

                // identify unassigned procedure-scoped declarations:
                unassignedDeclarations.AddRange(
                    locals.Where(local => assignments.All(a => (a.Scope.MemberName + a.Name) != (local.Scope.MemberName + local.Name)))
                          .Select(identifier => new VariableNotAssignedInspectionResult(Name, Severity, identifier.Context.ambiguousIdentifier(), result.QualifiedName)));

                // identify globals assigned in this module:
                assignedGlobals.AddRange(globals.Where(global => assignments.Any(a => a.Name == global.Global.GetText()))
                                                .Select(global => global.Global));
            }

            // identify unassigned globals:
            var assignedIdentifiers = assignedGlobals.Select(assigned => assigned.GetText());
            var unassignedGlobals = globals.Where(global => !assignedIdentifiers.Contains(global.Global.GetText()))
                                           .Select(identifier => new VariableNotAssignedInspectionResult(Name, Severity, identifier.Global, identifier.Name));
            unassignedDeclarations.AddRange(unassignedGlobals);

            return unassignedDeclarations;
        }
    }
}