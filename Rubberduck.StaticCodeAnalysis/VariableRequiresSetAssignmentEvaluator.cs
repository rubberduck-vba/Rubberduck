using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Inspections
{
    public static class VariableRequiresSetAssignmentEvaluator
    {
        public static IEnumerable<Declaration> GetDeclarationsPotentiallyRequiringSetAssignment(IEnumerable<Declaration> declarations)
        {
            //Reduce most of the declaration list with the easy ones
            var relevantDeclarations = declarations.Where(dec => dec.AsTypeName == Tokens.Variant
                                 || !SymbolList.ValueTypes.Contains(dec.AsTypeName)
                                 &&(MayRequireAssignmentUsingSet(dec) || RequiresAssignmentUsingSet(dec)));

            return relevantDeclarations;
        }

        /// <summary>
        /// Determines whether the 'Set' keyword needs to be added in the context of the specified identifier reference.
        /// </summary>
        /// <param name="reference">The identifier reference to analyze</param>
        /// <param name="state">The parser state</param>
        public static bool NeedsSetKeywordAdded(IdentifierReference reference, RubberduckParserState state)
        {
            var setStmtContext = reference.Context.GetAncestor<VBAParser.SetStmtContext>();
            return setStmtContext == null && RequiresSetAssignment(reference, state);
        }

        /// <summary>
        /// Determines whether the 'Set' keyword is required (whether it's present or not) for the specified identifier reference.
        /// </summary>
        /// <param name="reference">The identifier reference to analyze</param>
        /// <param name="state">The parser state</param>
        public static bool RequiresSetAssignment(IdentifierReference reference, RubberduckParserState state)
        {
            if (!reference.IsAssignment)
            {
                // reference isn't assigning its declaration; not interesting
                return false;
            }

            var setStmtContext = reference.Context.GetAncestor<VBAParser.SetStmtContext>();
            if (setStmtContext != null)
            {
                // don't assume Set keyword is legit...
                return reference.Declaration.IsObject;
            }

            var letStmtContext = reference.Context.GetAncestor<VBAParser.LetStmtContext>();
            if (letStmtContext == null)
            {
                // not an assignment
                return false;
            }

            var declaration = reference.Declaration;
            if (declaration.IsArray)
            {
                // arrays don't need a Set statement... todo figure out if array items are objects
                return false;
            }

            var isObjectVariable = declaration.IsObject;
            var isVariant = declaration.IsUndeclared || declaration.AsTypeName == Tokens.Variant;
            if (!isObjectVariable && !isVariant)
            {
                return false;
            }

            if (isObjectVariable)
            {
                // get the members of the returning type, a default member could make us lie otherwise
                var classModule = declaration.AsTypeDeclaration as ClassModuleDeclaration;
                if (classModule?.DefaultMember != null)
                {
                    var parameters = (classModule.DefaultMember as IParameterizedDeclaration)?.Parameters.ToArray() ?? Enumerable.Empty<ParameterDeclaration>().ToArray();
                    if (!parameters.Any() || parameters.All(p => p.IsOptional))
                    {
                        // assigned declaration has a default parameterless member, which is legally being assigned here.
                        // might be a good idea to flag that default member assignment though...
                        return false;
                    }
                }

                // assign declaration is an object without a default parameterless (or with all parameters optional) member - LHS needs a 'Set' keyword.
                return true;
            }

            // assigned declaration is a variant. we need to know about the RHS of the assignment.

            var expression = letStmtContext.expression();
            if (expression == null)
            {
                Debug.Assert(false, "RHS expression is empty? What's going on here?");
            }

            if (expression is VBAParser.NewExprContext)
            {
                // RHS expression is newing up an object reference - LHS needs a 'Set' keyword:
                return true;
            }

            var literalExpression = expression as VBAParser.LiteralExprContext;
            if (literalExpression?.literalExpression()?.literalIdentifier()?.objectLiteralIdentifier() != null)
            {
                // RHS is a 'Nothing' token - LHS needs a 'Set' keyword:
                return true;
            }

            // todo resolve expression return type

            var memberRefs = state.DeclarationFinder.IdentifierReferences(reference.ParentScoping.QualifiedName);
            var lastRef = memberRefs.LastOrDefault(r => !Equals(r, reference) && r.Context.GetAncestor<VBAParser.LetStmtContext>() == letStmtContext);
            if (lastRef?.Declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.ClassModule) ?? false)
            {
                // the last reference in the expression is referring to an object type
                return true;
            }
            if (lastRef?.Declaration.AsTypeName == Tokens.Object)
            {
                return true;
            }

            var accessibleDeclarations = state.DeclarationFinder.GetAccessibleDeclarations(reference.ParentScoping);
            foreach (var accessibleDeclaration in accessibleDeclarations.Where(d => d.IdentifierName == expression.GetText()))
            {
                if (accessibleDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule) || accessibleDeclaration.AsTypeName == Tokens.Object)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool MayRequireAssignmentUsingSet(Declaration declaration)
        {
            if (declaration.AsTypeName == Tokens.Variant)
            {
                return true;
            }

            if (declaration.IsArray)
            {
                return false;
            }

            if (declaration.AsTypeDeclaration != null)
            {
                if ((ClassModuleDeclaration.HasDefaultMember(declaration.AsTypeDeclaration)
                    || declaration.AsTypeDeclaration.DeclarationType == DeclarationType.Enumeration))
                {
                    return false;
                }
            }

            if (SymbolList.ValueTypes.Contains(declaration.AsTypeName))
            {
                return false;
            }
            return true;
        }

        private static bool RequiresAssignmentUsingSet(Declaration declaration)
        {
            if (declaration.AsTypeDeclaration != null)
            {
                return declaration.AsTypeDeclaration.DeclarationType == DeclarationType.ClassModule
                        && (((IsVariableOrParameter(declaration) 
                                && !declaration.IsSelfAssigned)
                            || (IsMemberWithReturnType(declaration)  
                                && declaration.IsTypeSpecified)));
            }
            return false;
        }

        private static bool IsMemberWithReturnType(Declaration item)
        {
            return (item.DeclarationType == DeclarationType.Function
                || item.DeclarationType == DeclarationType.PropertyGet);
        }

        private static bool IsVariableOrParameter(Declaration item)
        {
            return item.DeclarationType == DeclarationType.Variable
                    || item.DeclarationType == DeclarationType.Parameter;
        }
    }
}
