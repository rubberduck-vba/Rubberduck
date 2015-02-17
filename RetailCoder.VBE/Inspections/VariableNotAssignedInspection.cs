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
            var globals = FindGlobalVariables(parseResults).ToList();

            var assignedGlobals = new List<VBParser.AmbiguousIdentifierContext>();
            var unassignedDeclarations = new List<CodeInspectionResultBase>();

            foreach (var result in parseResults)
            {
                // module-scoped in this module:
                var declarations = GetModuleDeclarations(result, globals).ToList();
                var procedures = result.ParseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener()).ToList();

                // todo: replace anonymous types with actual types, and extract methods.

                // fetch & scope all assignments:
                var assignments = procedures.SelectMany(
                    procedure => procedure.GetContexts<VariableAssignmentListener, VBParser.AmbiguousIdentifierContext>(new VariableAssignmentListener())
                                         .Select(context => new
                                             {
                                                 Scope = new QualifiedMemberName(result.QualifiedName, ((dynamic)procedure).AmbiguousIdentifier().GetText()),
                                                 Name = context.GetText()
                                             }));

                // fetch & scope all procedure-scoped declarations:
                var locals = procedures.SelectMany(
                    procedure => procedure.GetContexts<DeclarationListener, ParserRuleContext>(new DeclarationListener())
                                          .OfType<VBParser.VariableSubStmtContext>()
                                          .Select(context => new
                                             {
                                                 Context = context,
                                                 Scope = new QualifiedMemberName(result.QualifiedName, ((dynamic)procedure).AmbiguousIdentifier().GetText()),
                                                 Name = context.AmbiguousIdentifier().GetText(),
                                                 Usages = procedure.GetContexts<VariableUsageListener, VBParser.AmbiguousIdentifierContext>(new VariableUsageListener())
                                                                   .Where(usage => usage.GetText() == context.AmbiguousIdentifier().GetText())
                                             })).ToList();

                // identify unassigned module-scoped declarations:
                unassignedDeclarations.AddRange(
                    declarations.Select(d => d.AmbiguousIdentifier())
                                .Where(d => globals.All(g => g.Context.GetText() != d.GetText()) 
                                        && assignments.All(a => a.Name != d.GetText()))
                                .Select(identifier => new VariableNotAssignedInspectionResult(Name, Severity, identifier, result.QualifiedName)));

                // identify unassigned procedure-scoped declarations:
                unassignedDeclarations.AddRange(
                    locals.Where(local => assignments.All(a => (a.Scope.MemberName + a.Name) != (local.Scope.MemberName + local.Name)))
                          .Select(identifier => new VariableNotAssignedInspectionResult(Name, Severity, identifier.Context.AmbiguousIdentifier(), result.QualifiedName)));

                // identify globals assigned in this module:
                assignedGlobals.AddRange(globals.Where(global => assignments.Any(a => a.Name == global.Context.GetText()))
                                                .Select(global => global.Context));

                // identify unused locals:
                unassignedDeclarations.AddRange(
                    locals.Where(local => !local.Usages.Any()).Select(local => new VariableNotUsedInspectionResult(Name, Severity, local.Context, local.Scope.ModuleScope)));
            }

            // identify unassigned globals:
            var assignedIdentifiers = assignedGlobals.Select(assigned => assigned.GetText());
            var unassignedGlobals = globals.Where(global => !assignedIdentifiers.Contains(global.Context.GetText()))
                                           .Select(identifier => new VariableNotAssignedInspectionResult(Name, Severity, identifier.Context, identifier.QualifiedName));
            unassignedDeclarations.AddRange(unassignedGlobals);

            return unassignedDeclarations;
        }

        private static IEnumerable<VBParser.VariableSubStmtContext> GetModuleDeclarations(VBComponentParseResult module, List<QualifiedContext<VBParser.AmbiguousIdentifierContext>> globals)
        {
            var declarations = 
                module.ParseTree.GetContexts<DeclarationSectionListener, ParserRuleContext>(new DeclarationSectionListener())
                    .OfType<VBParser.VariableSubStmtContext>()
                    .Where(variable => 
                        globals.All(global => 
                            !global.QualifiedName.Equals(module.QualifiedName) 
                         && !global.Context.GetText().Equals(variable.GetText())));

            return declarations;
        }

        private static IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> 
            FindGlobalVariables(IEnumerable<VBComponentParseResult> parseResults)
        {
            var globals = parseResults.SelectMany(
                result => result.ParseTree.GetContexts<DeclarationSectionListener, ParserRuleContext>(new DeclarationSectionListener())
                                .OfType<VBParser.VariableStmtContext>()
                                .Where(IsGlobal)
                                .SelectMany(context => 
                                    GetDeclaredIdentifiers(context).Select(variable => 
                                        variable.ToQualifiedContext(result.QualifiedName))));
            return globals;
        }

        private static bool IsGlobal(VBParser.VariableStmtContext context)
        {
            var visibility = context.Visibility();
            return visibility != null
                   && visibility.GetText() != Tokens.Private;

        }

        private static IEnumerable<VBParser.AmbiguousIdentifierContext> GetDeclaredIdentifiers(VBParser.VariableStmtContext context)
        {
            return context.VariableListStmt()
                          .VariableSubStmt()
                          .Select(variable => variable.AmbiguousIdentifier());
        }
    }
}