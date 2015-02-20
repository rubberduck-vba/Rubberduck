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
            var inspector = new IdentifierUsageInspector(parseResult);
            var issues = inspector.UnassignedGlobals()
                  .Union(inspector.UnassignedFields())
                  .Union(inspector.UnassignedLocals());

            foreach (var issue in issues)
            {
                yield return new VariableNotAssignedInspectionResult(Name, Severity, issue.Context, issue.QualifiedName);
            }
        }

        private static IEnumerable<VBParser.VariableSubStmtContext> 
            GetModuleDeclarations(VBComponentParseResult module, IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> globals)
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
            return context != null && context.GetText() != Tokens.Private;

        }

        private static IEnumerable<VBParser.AmbiguousIdentifierContext> GetDeclaredIdentifiers(VBParser.VariableStmtContext context)
        {
            return context.VariableListStmt()
                          .VariableSubStmt()
                          .Select(variable => variable.AmbiguousIdentifier());
        }
    }
}