using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class IdentifierUsageInspector
    {
        private readonly IEnumerable<VBComponentParseResult> _parseResult;
        private readonly IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _globals;
        private readonly IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _fields;
        private readonly IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _locals;
        private readonly IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _parameters;

        private readonly IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _assignments;
        private readonly IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _usages;

        public IdentifierUsageInspector(IEnumerable<VBComponentParseResult> parseResult)
        {
            _parseResult = parseResult;
            _globals = GetGlobals();
            _fields = GetFields(_globals);
            _locals = GetLocals();
            _parameters = GetParameters();
            _assignments = GetAssignments();
            _usages = GetIdentifierUsages(_assignments);
        }

        /// <summary>
        /// Gets all private fields whose name clashes with that of a global variable or constant.
        /// </summary>
        /// <remarks>
        /// VBA compiler resolves identifier references to the tightest applicable scope.
        /// </remarks>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> AmbiguousFieldNames()
        {
            foreach (var field in _fields)
            {
                var fieldName = field.Context.GetText();
                if (_globals.Any(global => global.Context.GetText() == fieldName))
                {
                    yield return field;
                }
            }
        }

        /// <summary>
        /// Gets all global-scope fields that are not assigned in any standard or class module.
        /// </summary>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnassignedGlobals()
        {
            var unassignedGlobals = _globals.Where(context => context.Context.Parent.GetType() != typeof(VBParser.ConstSubStmtContext))
                .Where(global => _assignments.Where(assignment => assignment.QualifiedName.Equals(global.QualifiedName))
                    .All(assignment => global.Context.GetText() != assignment.Context.GetText()));

            foreach (var unassignedGlobal in unassignedGlobals)
            {
                var global = unassignedGlobal;
                if (_assignments.Where(assignment => assignment.QualifiedName != global.QualifiedName)
                    .All(assignment => global.Context.GetText() != assignment.Context.GetText()))
                {
                    yield return global;
                }
            }
        }

        /// <summary>
        /// Gets all globals, fields and locals that are not assigned in their respective scope.
        /// </summary>
        /// <returns>
        /// Returns the declaration context's identifier.
        /// </returns>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> AllUnassignedVariables()
        {
            return UnassignedGlobals().Union(UnassignedFields().Union(UnassignedLocals()));
        }

        /// <summary>
        /// Gets all globals, fields and locals that are not used and not assigned in their respective scope.
        /// </summary>
        /// <returns>
        /// Returns the declaration context's identifier.
        /// </returns>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> AllUnusedVariables()
        {
            return UnusedGlobals().Union(UnusedFields().Union(UnusedLocals()));
        }

        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UndeclaredVariableUsages()
        {
            return _usages.Where(usage => _locals.All(local => local.MemberName != usage.MemberName || local.Context.GetText() != usage.Context.GetText())
                                        && _fields.All(field => field.QualifiedName != usage.QualifiedName || field.Context.GetText() != usage.Context.GetText())
                                        && _globals.All(global => global.Context.GetText() != usage.Context.GetText())
                                        && _parameters.All(parameter => parameter.MemberName != usage.MemberName || parameter.Context.GetText() != usage.Context.GetText()));
        }

        /// <summary>
        /// Gets all globals, fields and locals that are unassigned (used or not) in their respective scope.
        /// </summary>
        /// <returns>
        /// Returns the variable call context's identifier.
        /// </returns>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> AllUnassignedVariableUsages()
        {
            var variables = AllUnassignedVariables();
            return _usages.Where(usage => 
                variables.Any(variable => usage.QualifiedName == variable.QualifiedName
                                        && usage.Context.GetText() == variable.Context.GetText()));
        }

        /// <summary>
        /// Gets all module-scope fields that are not assigned.
        /// </summary>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnassignedFields()
        {
            var unassignedFields = _fields.Where(context => context.Context.Parent.GetType() != typeof(VBParser.ConstSubStmtContext))
                .Where(field => _assignments.Where(assignment => assignment.QualifiedName.Equals(field.QualifiedName))
                        .All(assignment => field.Context.GetText() != assignment.Context.GetText()));

            foreach (var field in unassignedFields)
            {
                yield return field;
            }
        }

        /// <summary>
        /// Gets all procedure-scope locals that are not assigned.
        /// </summary>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnassignedLocals()
        {
            var unassignedFields = _locals.Where(context => context.Context.Parent.GetType() != typeof(VBParser.ConstSubStmtContext))
                .Where(local => _assignments.Where(assignment => assignment.QualifiedName.Equals(local.QualifiedName))
                    .All(assignment => local.Context.GetText() != assignment.Context.GetText()));

            foreach (var field in unassignedFields)
            {
                yield return field;
            }
        }

        /// <summary>
        /// Gets all unassigned ByRef parameters.
        /// </summary>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnassignedByRefParameters()
        {
            var byRefParams = 
                (from parameter in _parameters
                let byRef = ((VBParser.ArgContext) parameter.Context.Parent).BYREF()
                let byVal = ((VBParser.ArgContext) parameter.Context.Parent).BYVAL()
                where byRef != null || (byRef == null && byVal == null)
                select parameter).ToList();

            var unassigned = _parameters.Where(parameter => byRefParams.Contains(parameter)
                && _assignments.Where(usage => usage.MemberName.Equals(parameter.MemberName))
                    .All(usage => parameter.Context.GetText() != usage.Context.GetText()));

            foreach (var context in unassigned)
            {
                yield return context;
            }
        }

        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnusedGlobals()
        {
            var unusedGlobals = _globals.Where(context => _usages.Where(usage => usage.QualifiedName.Equals(context.QualifiedName))
                    .All(usage => context.Context.GetText() != usage.Context.GetText()));

            foreach (var context in unusedGlobals)
            {
                yield return context;
            }
        }

        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnusedFields()
        {
            var unusedFields = _fields.Where(context =>
                _usages.Where(usage => usage.QualifiedName.Equals(context.QualifiedName))
                    .All(usage => context.Context.GetText() != usage.Context.GetText()));

            foreach (var context in unusedFields)
            {
                yield return context;
            }
        }

        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnusedLocals()
        {
            var unusedLocals = _locals.Where(context => 
                _usages.Where(usage => usage.MemberName.Equals(context.MemberName))
                    .All(usage => context.Context.GetText() != usage.Context.GetText()));

            foreach (var context in unusedLocals)
            {
                yield return context;
            }
        }

        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnusedParameters()
        {
            var unusedParameters = _parameters.Where(context =>
                _usages.Where(usage => usage.MemberName.Equals(context.MemberName))
                    .All(usage => context.Context.GetText() != usage.Context.GetText()));

            foreach (var context in unusedParameters)
            {
                yield return context;
            }
        }

        private IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetGlobals()
        {
            var result = new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();

            var modules = _parseResult.Where(e => e.Component.Type == vbext_ComponentType.vbext_ct_StdModule);
            foreach (var module in modules)
            {
                var listener = new DeclarationSectionListener(module.QualifiedName);
                var declarations = module.ParseTree
                    .GetContexts<DeclarationSectionListener, ParserRuleContext>(listener)
                    .ToList();

                result.AddRange(declarations.Select(declaration => declaration.Context)
                                            .OfType<VBParser.VariableStmtContext>()
                    .Where(declaration => IsGlobal(declaration.Visibility()))
                    .SelectMany(declaration => declaration.VariableListStmt().VariableSubStmt())
                    .Select(identifier => identifier.AmbiguousIdentifier().ToQualifiedContext(module.QualifiedName)));
            }

            return result;
        }

        private IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> 
            GetFields(IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> globals)
        {
            var result = new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            foreach (var module in _parseResult)
            {
                var listener = new DeclarationSectionListener(module.QualifiedName);
                var declarations = module.ParseTree
                    .GetContexts<DeclarationSectionListener, ParserRuleContext>(listener)
                    .Where(field => globals.All(global => global.QualifiedName.ModuleName == field.QualifiedName.ModuleName 
                                                       && global.Context.GetText() != field.Context.GetText()))
                    .ToList();

                result.AddRange(declarations.Select(declaration => declaration.Context)
                                            .OfType<VBParser.VariableSubStmtContext>()
                                            .Select(context => 
                        context.AmbiguousIdentifier().ToQualifiedContext(module.QualifiedName)));

                result.AddRange(declarations.Select(declaration => declaration.Context)
                                            .OfType<VBParser.TypeStmtContext>()
                                            .Select(context => 
                        context.AmbiguousIdentifier().ToQualifiedContext(module.QualifiedName)));
            }

            return result;
        }

        private IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetLocals()
        {
            var result = new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            foreach (var module in _parseResult)
            {
                var listener = new LocalDeclarationListener(module.QualifiedName);
                result.AddRange(module.ParseTree
                    .GetContexts<LocalDeclarationListener, VBParser.AmbiguousIdentifierContext>(listener));
            }

            return result;
        }

        private IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetParameters()
        {
            var result = new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            foreach (var module in _parseResult)
            {
                var listener = new ParameterListener(module.QualifiedName);
                result.AddRange(module.ParseTree
                    .GetContexts<ParameterListener, VBParser.AmbiguousIdentifierContext>(listener));
            }

            return result;
        }

        private IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetAssignments()
        {
            var result = new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            foreach (var module in _parseResult)
            {
                var listener = new VariableAssignmentListener(module.QualifiedName);
                result.AddRange(module.ParseTree
                    .GetContexts<VariableAssignmentListener, VBParser.AmbiguousIdentifierContext>(listener)
                    .Where(identifier => !IsConstant(identifier.Context) && !IsJoinedAssignemntDeclaration(identifier.Context)));
            }

            return result;
        }

        private IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetIdentifierUsages(IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> assignments)
        {
            var result = new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            foreach (var module in _parseResult)
            {
                var listener = new VariableReferencesListener(module.QualifiedName);

                // includes assignments
                var usages = module.ParseTree.GetContexts<VariableReferencesListener, VBParser.AmbiguousIdentifierContext>(listener);
                //result.AddRange(usages.Where(usage => assignments.Any(assignment => !usage.Equals(assignment))));
                result.AddRange(usages);
            }

            return result;
        }

        private static bool IsConstant(VBParser.AmbiguousIdentifierContext context)
        {
            return context.Parent.Parent.GetType() == typeof(VBParser.ConstSubStmtContext);
        }

        private static bool IsJoinedAssignemntDeclaration(VBParser.AmbiguousIdentifierContext context)
        {
            var declaration = context.Parent as VBParser.VariableSubStmtContext;
            if (declaration == null)
            {
                return false;
            }

            var asTypeClause = declaration.AsTypeClause();
            if (asTypeClause == null)
            {
                return false;
            }

            return asTypeClause.NEW() == null;
        }

        private static bool IsGlobal(VBParser.VisibilityContext context)
        {
            return context != null && context.GetText() != Tokens.Private;
        }
    }
}