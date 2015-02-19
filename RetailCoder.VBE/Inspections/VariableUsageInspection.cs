using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
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

            try
            {
                foreach (var result in parseResults)
                {
                    // module-scoped in this module:
                    var declarations = GetModuleDeclarations(result, globals).ToList();
                    var procedures = result.ParseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener(result.QualifiedName)).ToList();

                    // todo: replace anonymous types with actual types, and extract methods.

                    // fetch & scope all assignments:
                    var module = result;
                    var assignments = procedures.SelectMany(
                        procedure => procedure.Context.GetContexts<VariableAssignmentListener, VBParser.AmbiguousIdentifierContext>(new VariableAssignmentListener(result.QualifiedName))
                            .Select(context => new
                            {
                                Context = context,
                                Scope = new QualifiedMemberName(module.QualifiedName, ((dynamic)procedure.Context).AmbiguousIdentifier().GetText()),
                                Name = context.Context.GetText()
                            }));

                    // fetch & scope all procedure-scoped declarations:
                    var locals = procedures.SelectMany(
                        procedure => procedure.Context.GetContexts<DeclarationListener, ParserRuleContext>(new DeclarationListener(module.QualifiedName))
                            .OfType<VBParser.VariableSubStmtContext>()
                            .Select(context => new
                            {
                                Context = context,
                                Scope = new QualifiedMemberName(result.QualifiedName, ((dynamic)procedure.Context).AmbiguousIdentifier().GetText()),
                                Name = context.AmbiguousIdentifier().GetText(),
                                Usages = procedure.Context.GetContexts<VariableUsageListener, VBParser.AmbiguousIdentifierContext>(new VariableUsageListener(result.QualifiedName))
                                    .Where(usage => usage.Context.GetText() == context.AmbiguousIdentifier().GetText())
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

                    // identify assigned but unused locals:
                    // todo: figure out if it's possible to reuse this code AND to have a configurable VariableNotUsedInspection
                    unassignedDeclarations.AddRange(
                        locals.Where(local => local.Usages.All(usage => (usage.Context.Parent.Parent.Parent.Parent is VBParser.LetStmtContext)))
                            .Select(local => 
                                new VariableNotUsedInspectionResult(InspectionNames.VariableNotUsed, CodeInspectionSeverity.Error, local.Context, local.Scope.ModuleScope)));

                    // identify used but not declared locals:
                    unassignedDeclarations.AddRange(
                        assignments.Where(usage => declarations.All(declaration => declaration.AmbiguousIdentifier().GetText() != usage.Name))
                            .Select(usage => new VariableNotDeclaredInspectionResult(InspectionNames.VariableNotDeclared, CodeInspectionSeverity.Error, usage.Context.Context, result.QualifiedName)));
                }

            }
            catch (NullReferenceException e)
            {
                Console.WriteLine(e);
            }            // identify unassigned globals:
            var assignedIdentifiers = assignedGlobals.Select(assigned => assigned.GetText());
            var unassignedGlobals = globals.Where(global => !(global.Context.Parent is VBParser.TypeStmtContext) 
                                                         && !assignedIdentifiers.Contains(global.Context.GetText()))
                                           .Select(identifier => new VariableNotAssignedInspectionResult(Name, Severity, identifier.Context, identifier.QualifiedName));
            unassignedDeclarations.AddRange(unassignedGlobals);

            return unassignedDeclarations;
        }

        private static IEnumerable<VBParser.VariableSubStmtContext> GetModuleDeclarations(VBComponentParseResult module, List<QualifiedContext<VBParser.AmbiguousIdentifierContext>> globals)
        {
            var contexts = module.ParseTree
                .GetContexts<DeclarationSectionListener, ParserRuleContext>(new DeclarationSectionListener(module.QualifiedName))
                .ToList();

            var declarations =
                contexts.OfType<VBParser.VariableSubStmtContext>()
                    .Where(variable =>
                        globals.All(global =>
                            !global.QualifiedName.Equals(module.QualifiedName)
                            && !global.Context.GetText().Equals(variable.GetText())));

            return declarations;
        }

        private static IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> 
            FindGlobalVariables(IEnumerable<VBComponentParseResult> parseResults)
        {
            var globals = new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();

            foreach (var result in parseResults)
            {
                var module = result;
                var declarations = 
                    module.Context.GetContexts<DeclarationSectionListener, ParserRuleContext>(
                        new DeclarationSectionListener(result.QualifiedName)).ToList();
                
                globals.AddRange(declarations.OfType<VBParser.VariableStmtContext>()
                    .Where(declaration => IsGlobal(declaration.Visibility()))
                    .SelectMany(GetDeclaredIdentifiers)
                    .Select(identifier => identifier.ToQualifiedContext(module.QualifiedName)));

                globals.AddRange(declarations.OfType<VBParser.TypeStmtContext>()
                    .Where(declaration => IsGlobal(declaration.Visibility()))
                    .Select(declaration => declaration.AmbiguousIdentifier().ToQualifiedContext(module.QualifiedName)));
            }

            return globals;
        }

        private static bool IsGlobal(VBParser.VisibilityContext context)
        {
            return context != null
                   && context.GetText() != Tokens.Private;

        }

        private static IEnumerable<VBParser.AmbiguousIdentifierContext> GetDeclaredIdentifiers(VBParser.VariableStmtContext context)
        {
            return context.VariableListStmt()
                          .VariableSubStmt()
                          .Select(variable => variable.AmbiguousIdentifier());
        }
    }
}