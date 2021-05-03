using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Parsing.Binding
{
    public sealed class SimpleNameDefaultBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly ParserRuleContext _context;
        private readonly string _name;
        private readonly DeclarationType _propertySearchType;

        public SimpleNameDefaultBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            ParserRuleContext context,
            string name,
            StatementResolutionContext statementContext)
        {
            _declarationFinder = declarationFinder;
            _project = project;
            _module = module;
            _parent = parent;
            _context = context;
            // hack; SimpleNameContext.Identifier() excludes the square brackets
            _name = context.Start.Text == "[" && context.Stop.Text == "]" ? "[" + name + "]" : name;
            _propertySearchType = StatementContext.GetSearchDeclarationType(statementContext);
        }
        
        public bool IsPotentialLeftMatch { get; internal set; }

        public IBoundExpression Resolve()
        {
            IBoundExpression boundExpression = null;
            boundExpression = ResolveProcedureNamespace();
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveEnclosingModuleNamespace();
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveEnclosingProjectNamespace();
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveOtherProceduralModuleEnclosingProjectNamespace();
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveReferencedProjectNamespace();
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveModuleReferencedProjectNamespace();
            if (boundExpression != null)
            {
                return boundExpression;
            }

            if (_context.Start.Text == "[" && _context.Stop.Text == "]")
            {
                var bracketedExpression = _declarationFinder.OnBracketedExpression(_context.GetText(), _context, _module.QualifiedModuleName);
                return new SimpleNameExpression(bracketedExpression, ExpressionClassification.Unbound, _context);
            }
            
            var undeclaredLocal = _declarationFinder.OnUndeclaredVariable(_parent, _name, _context);
            return new SimpleNameExpression(undeclaredLocal, ExpressionClassification.Variable, _context);            
        }

        private IBoundExpression ResolveProcedureNamespace()
        {
            /*  Namespace tier 1:
                Procedure namespace: A local variable, reference parameter binding or constant whose implicit 
                or explicit definition precedes this expression in an enclosing procedure.                
            */
            if (!_parent.DeclarationType.HasFlag(DeclarationType.Function) && !_parent.DeclarationType.HasFlag(DeclarationType.Procedure))
            {
                return null;
            }
            var parameter = _declarationFinder.FindMemberEnclosingProcedure(_parent, _name, DeclarationType.Parameter);
            if (IsValidMatch(parameter, _name))
            {
                return new SimpleNameExpression(parameter, ExpressionClassification.Variable, _context);
            }
            var localVariable = _declarationFinder.FindMemberEnclosingProcedure(_parent, _name, DeclarationType.Variable)
                ?? _declarationFinder.FindMemberEnclosingProcedure(_parent, _name, DeclarationType.Variable)
                ;
            if (IsValidMatch(localVariable, _name))
            {
                return new SimpleNameExpression(localVariable, ExpressionClassification.Variable, _context);
            }
            var constant = _declarationFinder.FindMemberEnclosingProcedure(_parent, _name, DeclarationType.Constant);
            if (IsValidMatch(constant, _name))
            {
                return new SimpleNameExpression(constant, ExpressionClassification.Value, _context);
            }

            return null;
        }

        private IBoundExpression ResolveEnclosingModuleNamespace()
        {
            /*  Namespace tier 2:
                Enclosing Module namespace: A variable, constant, Enum type, Enum member, property, 
                function or subroutine defined at the module-level in the enclosing module.
            */
            var moduleVariable = _declarationFinder.FindMemberEnclosingModule(_module, _parent, _name, DeclarationType.Variable);
            if (IsValidMatch(moduleVariable, _name))
            {
                return new SimpleNameExpression(moduleVariable, ExpressionClassification.Variable, _context);
            }
            var moduleConstant = _declarationFinder.FindMemberEnclosingModule(_module, _parent, _name, DeclarationType.Constant);
            if (IsValidMatch(moduleConstant, _name))
            {
                return new SimpleNameExpression(moduleConstant, ExpressionClassification.Variable, _context);
            }
            var enumType = _declarationFinder.FindMemberEnclosingModule(_module, _parent, _name, DeclarationType.Enumeration);
            if (IsValidMatch(enumType, _name))
            {
                return new SimpleNameExpression(enumType, ExpressionClassification.Type, _context);
            }
            var enumMember = _declarationFinder.FindMemberEnclosingModule(_module, _parent, _name, DeclarationType.EnumerationMember);
            if (IsValidMatch(enumMember, _name))
            {
                return new SimpleNameExpression(enumMember, ExpressionClassification.Value, _context);
            }
            // Prioritize return value assignments over any other let/set property references.
            if (_parent.DeclarationType == DeclarationType.PropertyGet && _declarationFinder.IsMatch(_parent.IdentifierName, _name))
            {
                return new SimpleNameExpression(_parent, ExpressionClassification.Property, _context);
            }
            var property = _declarationFinder.FindMemberEnclosingModule(_module, _parent, _name, _propertySearchType);
            if (IsValidMatch(property, _name))
            {
                return new SimpleNameExpression(property, ExpressionClassification.Property, _context);
            }
            var function = _declarationFinder.FindMemberEnclosingModule(_module, _parent, _name, DeclarationType.Function);
            if (IsValidMatch(function, _name))
            {
                return new SimpleNameExpression(function, ExpressionClassification.Function, _context);
            }
            var subroutine = _declarationFinder.FindMemberEnclosingModule(_module, _parent, _name, DeclarationType.Procedure);
            if (IsValidMatch(subroutine, _name))
            {
                return new SimpleNameExpression(subroutine, ExpressionClassification.Subroutine, _context);
            }
            return null;
        }

        private IBoundExpression ResolveEnclosingProjectNamespace()
        {
            /*  Namespace tier 3:
                Enclosing Project namespace: The enclosing project itself, a referenced project, or a 
                procedural module contained in the enclosing project.
            */
            if (_declarationFinder.IsMatch(_project.ProjectName, _name))
            {
                return new SimpleNameExpression(_project, ExpressionClassification.Project, _context);
            }
            var referencedProject = _declarationFinder.FindReferencedProject(_project, _name);
            if (referencedProject != null)
            {
                return new SimpleNameExpression(referencedProject, ExpressionClassification.Project, _context);
            }
            if (_module.DeclarationType == DeclarationType.ProceduralModule && _declarationFinder.IsMatch(_module.IdentifierName, _name))
            {
                return new SimpleNameExpression(_module, ExpressionClassification.ProceduralModule, _context);
            }
            var proceduralModuleEnclosingProject = _declarationFinder.FindModuleEnclosingProjectWithoutEnclosingModule(_project, _module, _name, DeclarationType.ProceduralModule);
            if (proceduralModuleEnclosingProject != null)
            {
                return new SimpleNameExpression(proceduralModuleEnclosingProject, ExpressionClassification.ProceduralModule, _context);
            }
            var defaultInstanceVariableClass = _declarationFinder.FindDefaultInstanceVariableClassEnclosingProject(_project, _module, _name);
            if (defaultInstanceVariableClass != null)
            {
                return new SimpleNameExpression(defaultInstanceVariableClass, ExpressionClassification.Variable, _context);
            }
            return null;
        }

        private IBoundExpression ResolveOtherProceduralModuleEnclosingProjectNamespace()
        {
            /*  Namespace tier 4:
                Other Procedural Module in Enclosing Project namespace: An accessible variable, constant, 
                Enum type, Enum member, property, function or subroutine defined in a procedural module 
                within the enclosing project other than the enclosing module.  
            */
            var accessibleVariable = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, _name, DeclarationType.Variable);
            if (IsValidMatch(accessibleVariable, _name))
            {
                return new SimpleNameExpression(accessibleVariable, ExpressionClassification.Variable, _context);
            }
            var accessibleConstant = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, _name, DeclarationType.Constant);
            if (IsValidMatch(accessibleConstant, _name))
            {
                return new SimpleNameExpression(accessibleConstant, ExpressionClassification.Variable, _context);
            }
            var accessibleType = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, _name, DeclarationType.EnumerationMember);
            if (IsValidMatch(accessibleType, _name))
            {
                return new SimpleNameExpression(accessibleType, ExpressionClassification.Type, _context);
            }
            var accessibleMember = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, _name, DeclarationType.Enumeration);
            if (IsValidMatch(accessibleMember, _name))
            {
                return new SimpleNameExpression(accessibleMember, ExpressionClassification.Value, _context);
            }
            var accessibleProperty = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, _name, _propertySearchType);
            if (IsValidMatch(accessibleProperty, _name))
            {
                return new SimpleNameExpression(accessibleProperty, ExpressionClassification.Property, _context);
            }
            var accessibleFunction = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, _name, DeclarationType.Function);
            if (IsValidMatch(accessibleFunction, _name))
            {
                return new SimpleNameExpression(accessibleFunction, ExpressionClassification.Function, _context);
            }
            var accessibleSubroutine = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, _name, DeclarationType.Procedure);
            if (IsValidMatch(accessibleSubroutine, _name))
            {
                return new SimpleNameExpression(accessibleSubroutine, ExpressionClassification.Subroutine, _context);
            }
            return null;
        }

        private IBoundExpression ResolveReferencedProjectNamespace()
        {
            /*  Namespace tier 5:
                Referenced Project namespace: An accessible procedural module contained in a referenced 
                project.
            */
            var accessibleModule = _declarationFinder.FindModuleReferencedProject(_project, _module, _name, DeclarationType.ProceduralModule);
            if (accessibleModule != null)
            {
                return new SimpleNameExpression(accessibleModule, ExpressionClassification.ProceduralModule, _context);
            }
            var defaultInstanceVariableClass = _declarationFinder.FindDefaultInstanceVariableClassReferencedProject(_project, _module, _name);
            if (defaultInstanceVariableClass != null)
            {
                return new SimpleNameExpression(defaultInstanceVariableClass, ExpressionClassification.Type, _context);
            }
            return null;
        }

        private IBoundExpression ResolveModuleReferencedProjectNamespace()
        {
            /*  Namespace tier 6:
                Module in Referenced Project namespace: An accessible variable, constant, Enum type, 
                Enum member, property, function or subroutine defined in a procedural module or as a member 
                of the default instance of a global class module within a referenced project.  
            */

            // Part 1: Procedural module as parent
            var accessibleVariable = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, _name, DeclarationType.Variable);
            if (IsValidMatch(accessibleVariable, _name))
            {
                return new SimpleNameExpression(accessibleVariable, ExpressionClassification.Variable, _context);
            }
            var accessibleConstant = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, _name, DeclarationType.Constant);
            if (IsValidMatch(accessibleConstant, _name))
            {
                return new SimpleNameExpression(accessibleConstant, ExpressionClassification.Variable, _context);
            }
            var accessibleType = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, _name, DeclarationType.Enumeration);
            if (IsValidMatch(accessibleType, _name))
            {
                return new SimpleNameExpression(accessibleType, ExpressionClassification.Type, _context);
            }
            var accessibleMember = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, _name, DeclarationType.EnumerationMember);
            if (IsValidMatch(accessibleMember, _name))
            {
                return new SimpleNameExpression(accessibleMember, ExpressionClassification.Value, _context);
            }
            var accessibleProperty = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, _name, _propertySearchType);
            if (IsValidMatch(accessibleProperty, _name))
            {
                return new SimpleNameExpression(accessibleProperty, ExpressionClassification.Property, _context);
            }
            var accessibleFunction = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, _name, DeclarationType.Function);
            if (IsValidMatch(accessibleFunction, _name))
            {
                return new SimpleNameExpression(accessibleFunction, ExpressionClassification.Function, _context);
            }
            var accessibleSubroutine = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, _name, DeclarationType.Procedure);
            if (IsValidMatch(accessibleSubroutine, _name))
            {
                return new SimpleNameExpression(accessibleSubroutine, ExpressionClassification.Subroutine, _context);
            }

            // Part 2: Global class module as parent
            var globalClassModuleVariable = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, _name, DeclarationType.Variable);
            if (IsValidMatch(globalClassModuleVariable, _name))
            {
                return new SimpleNameExpression(globalClassModuleVariable, ExpressionClassification.Variable, _context);
            }
            var globalClassModuleConstant = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, _name, DeclarationType.Constant);
            if (IsValidMatch(globalClassModuleConstant, _name))
            {
                return new SimpleNameExpression(globalClassModuleConstant, ExpressionClassification.Variable, _context);
            }
            var globalClassModuleType = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, _name, DeclarationType.Enumeration);
            if (IsValidMatch(globalClassModuleType, _name))
            {
                return new SimpleNameExpression(globalClassModuleType, ExpressionClassification.Type, _context);
            }
            var globalClassModuleMember = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, _name, DeclarationType.EnumerationMember);
            if (IsValidMatch(globalClassModuleMember, _name))
            {
                return new SimpleNameExpression(globalClassModuleMember, ExpressionClassification.Value, _context);
            }
            var globalClassModuleProperty = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, _name, _propertySearchType);
            if (IsValidMatch(globalClassModuleProperty, _name))
            {
                return new SimpleNameExpression(globalClassModuleProperty, ExpressionClassification.Property, _context);
            }
            var globalClassModuleFunction = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, _name, DeclarationType.Function);
            if (IsValidMatch(globalClassModuleFunction, _name))
            {
                return new SimpleNameExpression(globalClassModuleFunction, ExpressionClassification.Function, _context);
            }
            var globalClassModuleSubroutine = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, _name, DeclarationType.Procedure);
            if (IsValidMatch(globalClassModuleSubroutine, _name))
            {
                return new SimpleNameExpression(globalClassModuleSubroutine, ExpressionClassification.Subroutine, _context);
            }

            return null;
        }

        private bool IsValidMatch(Declaration match, string name)
        {
            /*
               If the match has the name value "Left", references a function or subroutine that has no 
                parameters, or a property with a Property Get that has no parameters, the declared type of the 
                match is any type except a specific class, Object or Variant, and this simple name expression is 
                the <l-expression> within an index expression with an argument list containing 2 arguments, 
                discard the match and continue searching for a match on lower tiers. 

                Note: In other words, the built-in Left function is given highest priority?
            */
            if (match == null)
            {
                return false;
            }
            if (!IsPotentialLeftMatch || name.ToUpperInvariant() != "LEFT")
            {
                return true;
            }
            if (!IsFunctionSubroutinePropertyGet(match))
            {
                return true;
            }
            if (((IParameterizedDeclaration)match).Parameters.Any())
            {
                return true;
            }
            if (IsTypeDeclarationOfSpecificBaseType(match))
            {
                return false;
            }
            return true;
        }

            private static bool IsFunctionSubroutinePropertyGet(Declaration match)
            {
                return match.DeclarationType == DeclarationType.Function
                        || match.DeclarationType == DeclarationType.Procedure
                        || match.DeclarationType == DeclarationType.PropertyGet;
            }

            private static bool IsTypeDeclarationOfSpecificBaseType(Declaration match)
            {
                return match.AsTypeName != null
                        && match.AsTypeName.ToUpperInvariant() != "VARIANT"
                        && match.AsTypeName.ToUpperInvariant() != "OBJECT"
                        && match.AsTypeIsBaseType;
            }
    }
}
