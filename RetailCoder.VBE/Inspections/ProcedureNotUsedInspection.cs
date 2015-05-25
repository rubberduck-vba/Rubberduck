using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ProcedureNotUsedInspection : IInspection 
    {
        public ProcedureNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ProcedureNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var classes = parseResult.Declarations.Items.Where(item => !item.IsBuiltIn && item.DeclarationType == DeclarationType.Class).ToList();
            var modules = parseResult.Declarations.Items.Where(item => !item.IsBuiltIn && item.DeclarationType == DeclarationType.Module).ToList();

            var handlers = parseResult.Declarations.Items.Where(item => !item.IsBuiltIn && item.DeclarationType == DeclarationType.Control)
                .SelectMany(control => parseResult.Declarations.FindEventHandlers(control)).ToList();

            var withEventFields = parseResult.Declarations.Items.Where(item => !item.IsBuiltIn && item.DeclarationType == DeclarationType.Variable && item.IsWithEvents);
            handlers.AddRange(withEventFields.SelectMany(field => parseResult.Declarations.FindEventProcedures(field)));

            var forms = parseResult.Declarations.Items.Where(
                item => !item.IsBuiltIn && item.DeclarationType == DeclarationType.Class
                        && item.QualifiedName.QualifiedModuleName.Component.Type == vbext_ComponentType.vbext_ct_MSForm)
                .ToList();

            if (forms.Any())
            {
                handlers.AddRange(forms.SelectMany(form => parseResult.Declarations.FindFormEventHandlers(form)));
            }

            var issues = parseResult.Declarations.Items
                .Where(item => !item.IsBuiltIn && !IsIgnoredDeclaration(parseResult.Declarations, item, handlers, classes, modules))
                .Select(issue => new IdentifierNotUsedInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName.QualifiedModuleName));

            foreach (var item in DocumentEventHandlerPrefixes)
            {
                issues = issues.Where(issue => !issue.Name.Contains("'" + item));
            }

            return issues.ToList();
        }

        private static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function
        };

        private bool IsIgnoredDeclaration(Declarations declarations, Declaration declaration, IEnumerable<Declaration> handlers, IEnumerable<Declaration> classes, IEnumerable<Declaration> modules)
        {
            var enumerable = classes as IList<Declaration> ?? classes.ToList();
            var result = !ProcedureTypes.Contains(declaration.DeclarationType)
                || declaration.References.Any()
                || handlers.Contains(declaration)
                || IsPublicModuleMember(modules, declaration)
                || IsClassLifeCycleHandler(enumerable, declaration)
                || IsInterfaceMember(declarations, enumerable, declaration);

            return result;
        }

        /// <remarks>
        /// We cannot determine whether exposed members of standard modules are called or not,
        /// so we assume they are instead of flagging them as "never called".
        /// </remarks>
        private bool IsPublicModuleMember(IEnumerable<Declaration> modules, Declaration procedure)
        {
            if ((procedure.Accessibility != Accessibility.Implicit
                 && procedure.Accessibility != Accessibility.Public))
            {
                return false;
            }

            var parent = modules.Where(item => item.Project == procedure.Project)
                        .SingleOrDefault(item => item.IdentifierName == procedure.ComponentName);

            return parent != null;
        }

        private static readonly string[] DocumentEventHandlerPrefixes =
        {
            "Chart_",
            "Worksheet_",
            "Workbook_",
            "Document_",
            "Application_"
        };

        private static readonly string[] ClassLifeCycleHandlers =
        {
            "Class_Initialize",
            "Class_Terminate"
        };

        private bool IsClassLifeCycleHandler(IEnumerable<Declaration> classes, Declaration procedure)
        {
            if (!ClassLifeCycleHandlers.Contains(procedure.IdentifierName))
            {
                return false;
            }

            var parent = classes.Where(item => item.Project == procedure.Project)
                        .SingleOrDefault(item => item.IdentifierName == procedure.ComponentName);

            return parent != null;
        }

        /// <remarks>
        /// Interface implementation members are private, they're not called from an object
        /// variable reference of the type of the procedure's class, and whether they're called or not,
        /// they have to be implemented anyway, so removing them would break the code.
        /// Best just ignore them.
        /// </remarks>
        private bool IsInterfaceMember(Declarations declarations, IEnumerable<Declaration> classes, Declaration procedure)
        {
            // get the procedure's parent module
            var enumerable = classes as IList<Declaration> ?? classes.ToList();
            var parent = enumerable.Where(item => item.Project == procedure.Project)
                        .SingleOrDefault(item => item.IdentifierName == procedure.ComponentName);

            if (parent == null)
            {
                return false;
            }

            var interfaces = enumerable.Where(item => item.References.Any(reference =>
                    reference.Context.Parent is VBAParser.ImplementsStmtContext));

            if (interfaces.Select(i => i.ComponentName).Contains(procedure.ComponentName))
            {
                return true;
            }

            var result = GetImplementedInterfaceMembers(declarations, enumerable, procedure.ComponentName)
                .Contains(procedure.IdentifierName);

            return result;
        }

        private IEnumerable<string> GetImplementedInterfaceMembers(Declarations declarations, IEnumerable<Declaration> classes, string componentName)
        {
            var interfaces = classes.Where(item => item.References.Any(reference =>
                    reference.Context.Parent is VBAParser.ImplementsStmtContext
                    && reference.QualifiedModuleName.Component.Name == componentName));

            var members = interfaces.SelectMany(declarations.FindMembers)
                .Select(member => member.ComponentName + "_" + member.IdentifierName);
            return members;
        }
    }
}