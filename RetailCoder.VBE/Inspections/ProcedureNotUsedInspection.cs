using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ProcedureNotUsedInspection : IInspection 
    {
        public ProcedureNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ProcedureNotUsedInspection"; } }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return RubberduckUI.ProcedureNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var declarations = state.AllUserDeclarations.ToList();

            var classes = declarations.Where(item => item.DeclarationType == DeclarationType.Class).ToList();
            var modules = declarations.Where(item => item.DeclarationType == DeclarationType.Module).ToList();

            var handlers = declarations.Where(item => item.DeclarationType == DeclarationType.Control)
                .SelectMany(control => declarations.FindEventHandlers(control)).ToList();

            var withEventFields = declarations.Where(item => item.DeclarationType == DeclarationType.Variable && item.IsWithEvents);
            handlers.AddRange(withEventFields.SelectMany(field => declarations.FindEventProcedures(field)));

            var forms = declarations.Where(item => item.DeclarationType == DeclarationType.Class
                        && item.QualifiedName.QualifiedModuleName.Component.Type == vbext_ComponentType.vbext_ct_MSForm)
                .ToList();

            if (forms.Any())
            {
                handlers.AddRange(forms.SelectMany(form => declarations.FindFormEventHandlers(form)));
            }

            var issues = declarations
                .Where(item => !IsIgnoredDeclaration(declarations, item, handlers, classes, modules))
                .Select(issue => new IdentifierNotUsedInspectionResult(this, issue, issue.Context, issue.QualifiedName.QualifiedModuleName));

            issues = DocumentNames.DocumentEventHandlerPrefixes.Aggregate(issues, (current, item) => current.Where(issue => !issue.Name.Contains("'" + item)));

            return issues.ToList();
        }

        private static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function
        };

        private bool IsIgnoredDeclaration(IEnumerable<Declaration> declarations, Declaration declaration, IEnumerable<Declaration> handlers, IEnumerable<Declaration> classes, IEnumerable<Declaration> modules)
        {
            var enumerable = classes as IList<Declaration> ?? classes.ToList();
            var result = !ProcedureTypes.Contains(declaration.DeclarationType)
                || declaration.References.Any(r => !r.IsAssignment)
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
        private bool IsInterfaceMember(IEnumerable<Declaration> declarations, IEnumerable<Declaration> classes, Declaration procedure)
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

        private IEnumerable<string> GetImplementedInterfaceMembers(IEnumerable<Declaration> declarations, IEnumerable<Declaration> classes, string componentName)
        {
            var interfaces = classes.Where(item => item.References.Any(reference =>
                    reference.Context.Parent is VBAParser.ImplementsStmtContext
                    && reference.QualifiedModuleName.Component.Name == componentName));

            var members = interfaces.SelectMany(declarations.InScope)
                .Select(member => member.ComponentName + "_" + member.IdentifierName);
            return members;
        }
    }
}