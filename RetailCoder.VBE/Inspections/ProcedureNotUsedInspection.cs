using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ProcedureNotUsedInspection : IInspection /* note: deferred to v1.4 */
    {
        public ProcedureNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.ProcedureNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var handlers = parseResult.Declarations.Items.Where(item => item.DeclarationType == DeclarationType.Control)
                .SelectMany(control => parseResult.Declarations.FindEventHandlers(control))
                .ToList();

            var issues = parseResult.Declarations.Items
                .Where(item => !IsIgnoredDeclaration(parseResult.Declarations, item, handlers))
                .Select(issue => new IdentifierNotUsedInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName.QualifiedModuleName))
                .ToList();

            return issues;
        }

        private static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function
        };

        private bool IsIgnoredDeclaration(Declarations declarations, Declaration declaration, IEnumerable<Declaration> handlers)
        {
            var result = !ProcedureTypes.Contains(declaration.DeclarationType)
                || declaration.References.Any()
                || handlers.Contains(declaration)
                || IsPublicModuleMember(declarations, declaration)
                || IsClassLifeCycleHandler(declarations, declaration)
                || IsInterfaceMember(declarations, declaration);

            return result;
        }

        /// <remarks>
        /// We cannot determine whether exposed members of standard modules are called or not,
        /// so we assume they are instead of flagging them as "never called".
        /// </remarks>
        private bool IsPublicModuleMember(Declarations declarations, Declaration procedure)
        {
            var parent = declarations.Items.SingleOrDefault(item =>
                        item.Project == procedure.Project &&
                        item.IdentifierName == procedure.ComponentName && 
                        (item.DeclarationType == DeclarationType.Module));

            return parent != null && (procedure.Accessibility == Accessibility.Implicit
                                      || procedure.Accessibility == Accessibility.Public);
        }

        private static readonly string[] ClassLifeCycleHandlers =
        {
            "Class_Initialize",
            "Class_Terminate"
        };

        private bool IsClassLifeCycleHandler(Declarations declarations, Declaration procedure)
        {
            var parent = declarations.Items.SingleOrDefault(item =>
                        item.Project == procedure.Project && 
                        item.IdentifierName == procedure.ComponentName &&
                        (item.DeclarationType == DeclarationType.Class));

            return parent != null && ClassLifeCycleHandlers.Contains(procedure.IdentifierName);
        }

        /// <remarks>
        /// Interface implementation members are private, they're not called from an object
        /// variable reference of the type of the procedure's class, and whether they're called or not,
        /// they have to be implemented anyway, so removing them would break the code.
        /// Best just ignore them.
        /// </remarks>
        private bool IsInterfaceMember(Declarations declarations, Declaration procedure)
        {
            // get the procedure's parent module
            var parent = declarations.Items.SingleOrDefault(item =>
                        item.Project == procedure.Project && 
                        item.IdentifierName == procedure.ComponentName &&
                       (item.DeclarationType == DeclarationType.Class));

            if (parent == null)
            {
                return false;
            }

            var classes = declarations.Items.Where(item => item.DeclarationType == DeclarationType.Class);
            var interfaces = classes.Where(item => item.References.Any(reference =>
                    reference.Context.Parent is VBAParser.ImplementsStmtContext));

            if (interfaces.Select(i => i.ComponentName).Contains(procedure.ComponentName))
            {
                return true;
            }

            var result = GetImplementedInterfaceMembers(declarations, procedure.ComponentName)
                .Contains(procedure.IdentifierName);

            return result;
        }

        private IEnumerable<string> GetImplementedInterfaceMembers(Declarations declarations, string componentName)
        {
            var classes = declarations.Items.Where(item => item.DeclarationType == DeclarationType.Class);
            var interfaces = classes.Where(item => item.References.Any(reference =>
                    reference.Context.Parent is VBAParser.ImplementsStmtContext
                    && reference.QualifiedModuleName.ModuleName == componentName));

            var members = interfaces.SelectMany(declarations.FindMembers)
                .Select(member => member.ComponentName + "_" + member.IdentifierName);
            return members;
        }
    }
}