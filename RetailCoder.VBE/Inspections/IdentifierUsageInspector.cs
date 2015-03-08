using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
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
        private readonly HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _globals;
        private readonly HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _fields;
        private readonly HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _locals;
        private readonly HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _parameters;

        private readonly HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _assignments;
        private readonly HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _usages;

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
            // not used.... yet.

            foreach (var field in _fields)
            {
                var fieldName = field.Context.GetText();
                if (_globals.Any(global => global.Context.GetText() == fieldName))
                {
                    yield return field;
                }
            }
        }


        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _unassignedGlobals;
        /// <summary>
        /// Gets all global-scope fields that are not assigned in any standard or class module.
        /// </summary>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnassignedGlobals()
        {
            if (_unassignedGlobals != null)
            {
                return _unassignedGlobals;
            }

            _unassignedGlobals = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();

            var unassignedGlobals = _globals.Where(context => context.Context.Parent.GetType() != typeof(VBParser.ConstSubStmtContext))
                .Where(global => _assignments.Where(assignment => assignment.QualifiedName.Equals(global.QualifiedName))
                    .All(assignment => global.Context.GetText() != assignment.Context.GetText()));

            foreach (var unassignedGlobal in unassignedGlobals)
            {
                var global = unassignedGlobal;
                if (_assignments.Where(assignment => assignment.QualifiedName != global.QualifiedName)
                    .All(assignment => global.Context.GetText() != assignment.Context.GetText()))
                {
                    _unassignedGlobals.Add(global);
                }
            }

            return _unassignedGlobals;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _allUnassignedVariables;
        /// <summary>
        /// Gets all globals, fields and locals that are not assigned in their respective scope.
        /// </summary>
        /// <returns>
        /// Returns the declaration context's identifier.
        /// </returns>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> AllUnassignedVariables()
        {
            if (_allUnassignedVariables != null)
            {
                return _allUnassignedVariables;
            }

            _allUnassignedVariables = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(UnassignedGlobals().Union(UnassignedFields().Union(UnassignedLocals())));
            return _allUnassignedVariables;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _allUnusedVariables;
        /// <summary>
        /// Gets all globals, fields and locals that are not used and not assigned in their respective scope.
        /// </summary>
        /// <returns>
        /// Returns the declaration context's identifier.
        /// </returns>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> AllUnusedVariables()
        {
            if (_allUnusedVariables != null)
            {
                return _allUnusedVariables;
            }

            _allUnusedVariables = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(UnusedGlobals().Union(UnusedFields().Union(UnusedLocals())));
            return _allUnusedVariables;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _undeclaredVariableUsages;
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UndeclaredVariableUsages()
        {
            if (_undeclaredVariableUsages != null)
            {
                return _undeclaredVariableUsages;
            }

            var undeclared = _usages.Where(usage => _locals.All(local => local.MemberName != usage.MemberName && local.Context.GetText() != usage.Context.GetText())
                                        && _fields.All(field => field.QualifiedName != usage.QualifiedName && field.Context.GetText() != usage.Context.GetText())
                                        && _globals.All(global => global.Context.GetText() != usage.Context.GetText())
                                        && _parameters.All(parameter => parameter.MemberName != usage.MemberName && parameter.Context.GetText() != usage.Context.GetText()));

            _undeclaredVariableUsages = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(undeclared);
            return _undeclaredVariableUsages;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _allUnassignedVariableUsages;
        /// <summary>
        /// Gets all globals, fields and locals that are unassigned (used or not) in their respective scope.
        /// </summary>
        /// <returns>
        /// Returns the variable call context's identifier.
        /// </returns>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> AllUnassignedVariableUsages()
        {
            if (_allUnassignedVariableUsages != null)
            {
                return _allUnassignedVariableUsages;
            }

            var variables = AllUnassignedVariables();
            _allUnassignedVariableUsages = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(
                                                              _usages.Where(usage => variables.Any(variable => usage.QualifiedName == variable.QualifiedName
                                                              && usage.Context.GetText() == variable.Context.GetText()))
                                                              );

            return _allUnassignedVariableUsages;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _unassignedFields;
        /// <summary>
        /// Gets all module-scope fields that are not assigned.
        /// </summary>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnassignedFields()
        {
            if (_unassignedFields != null)
            {
                return _unassignedFields;
            }

            var userTypes = _fields.Select(field => field.Context.Parent)
                .OfType<VBParser.TypeStmtContext>()
                .Select(t => t.AmbiguousIdentifier());

            var userTypeFields = _fields.Select(field => field.Context.Parent)
                .OfType<VBParser.VariableSubStmtContext>()
                .Where(v => v.AsTypeClause() != null
                            && userTypes.Any(udt => udt.GetText() == v.AsTypeClause().Type().GetText()))
                .Select(v => v.AmbiguousIdentifier());

            _unassignedFields = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(
                            _fields.Where(context => userTypeFields.All(f => f.GetText() != context.Context.GetText()) // note: weak
                                && context.Context.Parent.GetType() != typeof(VBParser.ConstSubStmtContext))
                            .Where(field => _assignments.Where(assignment => assignment.QualifiedName.Equals(field.QualifiedName))
                                    .All(assignment => field.Context.GetText() != assignment.Context.GetText()))
                            );

            return _unassignedFields;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _unassignedLocals;
        /// <summary>
        /// Gets all procedure-scope locals that are not assigned.
        /// </summary>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnassignedLocals()
        {
            if (_unassignedLocals != null)
            {
                return _unassignedLocals;
            }

            _unassignedLocals = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(
                _locals.Where(context => context.Context.Parent.GetType() != typeof(VBParser.ConstSubStmtContext))
                .Where(local => _assignments.Where(assignment => assignment.QualifiedName.Equals(local.QualifiedName))
                    .All(assignment => local.Context.GetText() != assignment.Context.GetText()))
                );

            return _unassignedLocals;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _unassignedByRefParameters;
        /// <summary>
        /// Gets all unassigned ByRef parameters.
        /// </summary>
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnassignedByRefParameters()
        {
            if (_unassignedByRefParameters != null)
            {
                return _unassignedByRefParameters;
            }

            var byRefParams = 
                (from parameter in _parameters
                let byRef = ((VBParser.ArgContext) parameter.Context.Parent).BYREF()
                let byVal = ((VBParser.ArgContext) parameter.Context.Parent).BYVAL()
                where byRef != null || (byRef == null && byVal == null)
                select parameter).ToList();

            _unassignedByRefParameters = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(
                _parameters.Where(parameter => byRefParams.Contains(parameter)
                && _assignments.Where(usage => usage.MemberName.Equals(parameter.MemberName))
                    .All(usage => parameter.Context.GetText() != usage.Context.GetText()))
                );

            return _unassignedByRefParameters;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _unusedGlobals;
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnusedGlobals()
        {
            if (_unusedGlobals != null)
            {
                return _unusedGlobals;
            }

            _unusedGlobals = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(
                    _globals.Where(context => _usages.Where(usage => usage.Context.GetText() == context.Context.GetText())
                    .All(usage => context.Context.GetText() != usage.Context.GetText()))
                    );

            return _unusedGlobals;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _unusedFields;
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnusedFields()
        {
            if (_unusedFields != null)
            {
                return _unusedFields;
            }

            _unusedFields = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(
                _fields.Where(context => _globals.All(global => global != context) &&
                _usages.Where(usage => usage.QualifiedName.Equals(context.QualifiedName))
                    .All(usage => context.Context.GetText() != usage.Context.GetText()))
               );

            return _unusedFields;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _unusedLocals;
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnusedLocals()
        {
            if (_unusedLocals != null)
            {
                return _unusedLocals;
            }

            _unusedLocals = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(
                _locals.Where(context => 
                _usages.Where(usage => usage.MemberName.Equals(context.MemberName))
                    .All(usage => context.Context.GetText() != usage.Context.GetText()))
                );
            return _unusedLocals;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> _unusedParameters;
        public IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> UnusedParameters()
        {
            if (_unusedParameters != null)
            {
                return _unusedParameters;
            }

            _unusedParameters = new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(
                _parameters.Where(context => _usages
                    .Where(usage => usage.QualifiedName == context.QualifiedName).All(usage => context.Context.GetText() != usage.Context.GetText()))
                );
            return _unusedParameters;
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetGlobals()
        {
            var result = new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();

            var modules = _parseResult.Where(e => e.Component.Type == vbext_ComponentType.vbext_ct_StdModule);
            foreach (var module in modules)
            {
                var scope = module;
                var listener = new DeclarationSectionListener(module.QualifiedName);
                var declarations = module.ParseTree
                    .GetContexts<DeclarationSectionListener, ParserRuleContext>(listener)
                    .Select(declaration => declaration.Context)
                                            .OfType<VBParser.VariableStmtContext>()
                    .Where(declaration => IsGlobal(declaration.Visibility()))
                    .SelectMany(declaration => declaration.VariableListStmt().VariableSubStmt())
                    .Select(identifier => identifier.AmbiguousIdentifier().ToQualifiedContext(scope.QualifiedName));

                result.AddRange(declarations);
            }

            return new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> 
            GetFields(IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> globals)
        {
            var result = new List<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            foreach (var module in _parseResult)
            {
                var listener = new DeclarationSectionListener(module.QualifiedName);
                var declarations = module.ParseTree
                    .GetContexts<DeclarationSectionListener, ParserRuleContext>(listener)
                    .Where(field => globals.All(global => global.QualifiedName.ModuleName != field.QualifiedName.ModuleName 
                                                       && global.Context.GetText() != field.Context.GetText()))
                    .ToList();

                result.AddRange(declarations.Select(declaration => declaration.Context)
                                            .OfType<VBParser.VariableSubStmtContext>()
                                            .Select(context => 
                        context.AmbiguousIdentifier().ToQualifiedContext(module.QualifiedName)));
            }

            return new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetLocals()
        {
            var result = new ConcurrentBag<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            Parallel.ForEach(_parseResult, module =>
            {
                var listener = new LocalDeclarationListener(module.QualifiedName);
                foreach (var local in module.ParseTree
                    .GetContexts<LocalDeclarationListener, VBParser.AmbiguousIdentifierContext>(listener))
                {
                    result.Add(local);
                }
            });

            return new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetParameters()
        {
            var result = new ConcurrentBag<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            Parallel.ForEach(_parseResult, module =>
            {
                var listener = new ParameterListener(module.QualifiedName);
                foreach (var parameter in module.ParseTree
                    .GetContexts<ParameterListener, VBParser.AmbiguousIdentifierContext>(listener))
                {
                    result.Add(parameter);
                }
            });

            return new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetAssignments()
        {
            var result = new ConcurrentBag<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            Parallel.ForEach(_parseResult, module =>
            {
                var listener = new VariableAssignmentListener(module.QualifiedName);
                foreach (var assignment in module.ParseTree
                    .GetContexts<VariableAssignmentListener, VBParser.AmbiguousIdentifierContext>(listener)
                    .Where(identifier => !IsConstant(identifier.Context) && !IsJoinedAssignemntDeclaration(identifier.Context)))
                {
                    result.Add(assignment);
                }
            });

            return new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>> GetIdentifierUsages(IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> assignments)
        {
            if (_usages != null)
            {
                return _usages;
            }

            var result = new ConcurrentBag<QualifiedContext<VBParser.AmbiguousIdentifierContext>>();
            Parallel.ForEach(_parseResult, module =>
            {
                var listener = new VariableReferencesListener(module.QualifiedName);

                var usages = module.ParseTree.GetContexts<VariableReferencesListener, VBParser.AmbiguousIdentifierContext>(listener);
                foreach (var usage in usages.Where(usage => !IsAssignmentUsage(usage, assignments)))
                {
                    result.Add(usage);
                }
            });

            return new HashSet<QualifiedContext<VBParser.AmbiguousIdentifierContext>>(result);
        }

        private bool IsAssignmentUsage(QualifiedContext<VBParser.AmbiguousIdentifierContext> usage,
            IEnumerable<QualifiedContext<VBParser.AmbiguousIdentifierContext>> assignments)
        {
            return assignments.Any(assignment => assignment == usage);
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