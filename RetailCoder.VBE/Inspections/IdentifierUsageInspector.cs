using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class IdentifierUsageInspector
    {
        private readonly IEnumerable<VBComponentParseResult> _parseResult;
        private readonly HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _globals;
        private readonly HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _fields;
        private readonly HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _locals;
        private readonly HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _parameters;
        private readonly HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _userDefinedTypes = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();

        private readonly HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _assignments;
        private readonly HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _usages;

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
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> AmbiguousFieldNames()
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


        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _unassignedGlobals;
        /// <summary>
        /// Gets all global-scope fields that are not assigned in any standard or class module.
        /// </summary>
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> UnassignedGlobals()
        {
            if (_unassignedGlobals != null)
            {
                return _unassignedGlobals;
            }

            _unassignedGlobals = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();

            var unassignedGlobals = _globals.Where(context => context.Context.Parent.GetType() != typeof(VBAParser.ConstSubStmtContext))
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

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _allUnassignedVariables;
        /// <summary>
        /// Gets all globals, fields and locals that are not assigned in their respective scope.
        /// </summary>
        /// <returns>
        /// Returns the declaration context's identifier.
        /// </returns>
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> AllUnassignedVariables()
        {
            if (_allUnassignedVariables != null)
            {
                return _allUnassignedVariables;
            }

            _allUnassignedVariables = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(UnassignedGlobals().Union(UnassignedFields().Union(UnassignedLocals())));
            return _allUnassignedVariables;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _allUnusedVariables;
        /// <summary>
        /// Gets all globals, fields and locals that are not used and not assigned in their respective scope.
        /// </summary>
        /// <returns>
        /// Returns the declaration context's identifier.
        /// </returns>
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> AllUnusedVariables()
        {
            if (_allUnusedVariables != null)
            {
                return _allUnusedVariables;
            }

            _allUnusedVariables = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(UnusedGlobals().Union(UnusedFields().Union(UnusedLocals())));
            return _allUnusedVariables;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _undeclaredVariableUsages;
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> UndeclaredVariableUsages()
        {
            if (_undeclaredVariableUsages != null)
            {
                return _undeclaredVariableUsages;
            }

            var undeclared = _usages.Where(usage => _locals.All(local => local.MemberName != usage.MemberName && local.Context.GetText() != usage.Context.GetText())
                                        && _fields.All(field => field.QualifiedName != usage.QualifiedName && field.Context.GetText() != usage.Context.GetText())
                                        && _globals.All(global => global.Context.GetText() != usage.Context.GetText())
                                        && _parameters.All(parameter => parameter.MemberName != usage.MemberName && parameter.Context.GetText() != usage.Context.GetText()));

            _undeclaredVariableUsages = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(undeclared);
            return _undeclaredVariableUsages;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _allUnassignedVariableUsages;
        /// <summary>
        /// Gets all globals, fields and locals that are unassigned (used or not) in their respective scope.
        /// </summary>
        /// <returns>
        /// Returns the variable call context's identifier.
        /// </returns>
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> AllUnassignedVariableUsages()
        {
            if (_allUnassignedVariableUsages != null)
            {
                return _allUnassignedVariableUsages;
            }

            var variables = AllUnassignedVariables();
            _allUnassignedVariableUsages = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(
                                                              _usages.Where(usage => variables.Any(variable => usage.QualifiedName == variable.QualifiedName
                                                              && usage.Context.GetText() == variable.Context.GetText()))
                                                              );

            return _allUnassignedVariableUsages;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _unassignedFields;
        /// <summary>
        /// Gets all module-scope fields that are not assigned.
        /// </summary>
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> UnassignedFields()
        {
            if (_unassignedFields != null)
            {
                return _unassignedFields;
            }

            _unassignedFields = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(
                            _fields.Where(context => !IsUserDefinedType(context.Context.Parent as VBAParser.VariableSubStmtContext) 
                                && context.Context.Parent.GetType() != typeof(VBAParser.ConstSubStmtContext))
                            .Where(field => _assignments.Where(assignment => assignment.QualifiedName.Equals(field.QualifiedName))
                                    .All(assignment => field.Context.GetText() != assignment.Context.GetText()))
                            );

            return _unassignedFields;
        }

        private bool IsUserDefinedType(VBAParser.VariableSubStmtContext context)
        {
            if (context == null)
            {
                return false;
            }

            var type = context.asTypeClause() == null
                ? string.Empty
                : context.asTypeClause().type().GetText();

            // note: scoping issue; a private type could be used in another module (but that wouldn't compile)
            // Rubberduck inspections assume VBA code compiles though.
            return _userDefinedTypes.Any(udt => udt.Context.GetText() == type);
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _unassignedLocals;
        /// <summary>
        /// Gets all procedure-scope locals that are not assigned.
        /// </summary>
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> UnassignedLocals()
        {
            if (_unassignedLocals != null)
            {
                return _unassignedLocals;
            }

            _unassignedLocals = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(
                _locals.Where(context => context.Context.Parent.GetType() != typeof(VBAParser.ConstSubStmtContext))
                .Where(local => _assignments.Where(assignment => assignment.QualifiedName.Equals(local.QualifiedName))
                    .All(assignment => local.Context.GetText() != assignment.Context.GetText()))
                );

            return _unassignedLocals;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _unassignedByRefParameters;
        /// <summary>
        /// Gets all unassigned ByRef parameters.
        /// </summary>
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> UnassignedByRefParameters()
        {
            if (_unassignedByRefParameters != null)
            {
                return _unassignedByRefParameters;
            }

            var byRefParams = 
                (from parameter in _parameters
                let byRef = ((VBAParser.ArgContext) parameter.Context.Parent).BYREF()
                let byVal = ((VBAParser.ArgContext) parameter.Context.Parent).BYVAL()
                where byRef != null || (byRef == null && byVal == null)
                select parameter).ToList();

            _unassignedByRefParameters = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(
                _parameters.Where(parameter => byRefParams.Contains(parameter)
                && _assignments.Where(usage => usage.MemberName.Equals(parameter.MemberName))
                    .All(usage => parameter.Context.GetText() != usage.Context.GetText()))
                );

            return _unassignedByRefParameters;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _unusedGlobals;
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> UnusedGlobals()
        {
            if (_unusedGlobals != null)
            {
                return _unusedGlobals;
            }

            _unusedGlobals = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(
                    _globals.Where(context => _usages.Where(usage => usage.Context.GetText() == context.Context.GetText())
                    .All(usage => context.Context.GetText() != usage.Context.GetText()))
                    );

            return _unusedGlobals;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _unusedFields;
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> UnusedFields()
        {
            if (_unusedFields != null)
            {
                return _unusedFields;
            }

            _unusedFields = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(
                _fields.Where(context => _globals.All(global => global != context) &&
                _usages.Where(usage => usage.QualifiedName.Equals(context.QualifiedName))
                    .All(usage => context.Context.GetText() != usage.Context.GetText()))
               );

            return _unusedFields;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _unusedLocals;
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> UnusedLocals()
        {
            if (_unusedLocals != null)
            {
                return _unusedLocals;
            }

            _unusedLocals = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(
                _locals.Where(context => 
                _usages.Where(usage => usage.MemberName.Equals(context.MemberName))
                    .All(usage => context.Context.GetText() != usage.Context.GetText()))
                );
            return _unusedLocals;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _unusedParameters;
        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> UnusedParameters()
        {
            if (_unusedParameters != null)
            {
                return _unusedParameters;
            }

            _unusedParameters = new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(
                _parameters.Where(context => _usages
                    .Where(usage => usage.QualifiedName == context.QualifiedName).All(usage => context.Context.GetText() != usage.Context.GetText()))
                );
            return _unusedParameters;
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> GetGlobals()
        {
            var result = new List<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();

            var modules = _parseResult.Where(e => e.Component.Type == vbext_ComponentType.vbext_ct_StdModule);
            foreach (var module in modules)
            {
                var scope = module;
                var listener = new DeclarationSectionListener(module.QualifiedName);
                var declarations = module.ParseTree
                    .GetContexts<DeclarationSectionListener, ParserRuleContext>(listener)
                    .Select(declaration => declaration.Context).OfType<VBAParser.VariableStmtContext>()
                    .Where(declaration => IsGlobal(declaration.visibility()))
                    .SelectMany(declaration => declaration.variableListStmt().variableSubStmt())
                    .Select(identifier => identifier.ambiguousIdentifier().ToQualifiedContext(scope.QualifiedName))
                    .ToList();

                result.AddRange(declarations);
            }

            return new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> 
            GetFields(IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> globals)
        {
            var result = new List<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();
            foreach (var module in _parseResult)
            {
                var listener = new DeclarationSectionListener(module.QualifiedName);
                var declarations = module.ParseTree
                    .GetContexts<DeclarationSectionListener, ParserRuleContext>(listener)
                    .Where(field => globals.All(global => global.QualifiedName.ModuleName != field.QualifiedName.ModuleName 
                                                       && global.Context.GetText() != field.Context.GetText()))
                    .ToList();

                foreach (var udt in declarations.Where(declaration => declaration.Context.Parent is VBAParser.TypeStmtContext))
                {
                    _userDefinedTypes.Add(((VBAParser.TypeStmtContext)udt.Context.Parent).ambiguousIdentifier().ToQualifiedContext(module.QualifiedName));
                }

                result.AddRange(declarations.Select(declaration => declaration.Context)
                                            .OfType<VBAParser.VariableSubStmtContext>()
                                            .Select(context => 
                        context.ambiguousIdentifier().ToQualifiedContext(module.QualifiedName)));
            }

            return new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> GetLocals()
        {
            var result = new ConcurrentBag<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();
            Parallel.ForEach(_parseResult, module =>
            {
                var listener = new LocalDeclarationListener(module.QualifiedName);
                foreach (var local in module.ParseTree
                    .GetContexts<LocalDeclarationListener, VBAParser.AmbiguousIdentifierContext>(listener))
                {
                    result.Add(local);
                }
            });

            return new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> GetParameters()
        {
            var result = new ConcurrentBag<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();
            Parallel.ForEach(_parseResult, module =>
            {
                var listener = new ParameterListener(module.QualifiedName);
                foreach (var parameter in module.ParseTree
                    .GetContexts<ParameterListener, VBAParser.AmbiguousIdentifierContext>(listener))
                {
                    result.Add(parameter);
                }
            });

            return new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> GetAssignments()
        {
            var result = new ConcurrentBag<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();
            Parallel.ForEach(_parseResult, module =>
            {
                var listener = new VariableAssignmentListener(module.QualifiedName);
                foreach (var assignment in module.ParseTree
                    .GetContexts<VariableAssignmentListener, VBAParser.AmbiguousIdentifierContext>(listener)
                    .Where(identifier => !IsConstant(identifier.Context) && !IsJoinedAssignemntDeclaration(identifier.Context)))
                {
                    result.Add(assignment);
                }
            });

            return new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(result);
        }

        private HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> GetIdentifierUsages(IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> assignments)
        {
            if (_usages != null)
            {
                return _usages;
            }

            var result = new ConcurrentBag<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();
            Parallel.ForEach(_parseResult, module =>
            {
                var listener = new VariableReferencesListener(module.QualifiedName);

                var usages = module.ParseTree.GetContexts<VariableReferencesListener, VBAParser.AmbiguousIdentifierContext>(listener);
                foreach (var usage in usages.Where(usage => !IsAssignmentUsage(usage, assignments)))
                {
                    result.Add(usage);
                }
            });

            return new HashSet<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>(result);
        }

        private bool IsAssignmentUsage(QualifiedContext<VBAParser.AmbiguousIdentifierContext> usage,
            IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> assignments)
        {
            return assignments.Any(assignment => assignment == usage);
        }

        private static bool IsConstant(VBAParser.AmbiguousIdentifierContext context)
        {
            return context.Parent.Parent.GetType() == typeof(VBAParser.ConstSubStmtContext);
        }

        private static bool IsJoinedAssignemntDeclaration(VBAParser.AmbiguousIdentifierContext context)
        {
            var declaration = context.Parent as VBAParser.VariableSubStmtContext;
            if (declaration == null)
            {
                return false;
            }

            var asTypeClause = declaration.asTypeClause();
            if (asTypeClause == null)
            {
                return false;
            }

            return asTypeClause.NEW() == null;
        }

        private static bool IsGlobal(VBAParser.VisibilityContext context)
        {
            return context != null && context.GetText() != Tokens.Private;
        }
    }
}