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
            // reference isn't assigning its declaration
            if (!reference.IsAssignment) { return false; }

            // Set keyword is already there.
            // Returning reference.Declaration.IsObject allows flagging redundant Set keyword!
            var setStmtContext = reference.Context.GetAncestor<VBAParser.SetStmtContext>();
            if (setStmtContext != null) { return reference.Declaration.IsObject; }

            // Temporal coupling: LetStmtContext wouldn't be present given a SetStmtContext.
            // If both SetStmtContext and LetStmtContext are missing, we're not looking at an assignment expression.
            var letStmtContext = reference.Context.GetAncestor<VBAParser.LetStmtContext>();
            if (letStmtContext == null) { return false; }

            var declaration = reference.Declaration;
            var isObjectVariable = declaration.IsObject;
            if (declaration.IsArray)
            {
                // this is an array of object types (explicitly declared as such)
                return isObjectVariable;
            }

            // at this point we need to know if what we're looking at is an object or a variant.
            // if it's neither, we're done here.
            // if assigned declaration is an object with no default parameterless member, Set keyword is required (and missing!).
            if (isObjectVariable)
            {
                var classModule = declaration.AsTypeDeclaration as ClassModuleDeclaration;
                return !HasParameterlessDefaultMember(classModule);
            }

            // at this point if we're not looking at a variant, we can't say Set keyword is required/missing.
            var isVariant = declaration.IsUndeclared || declaration.AsTypeName == Tokens.Variant;
            if (!isVariant) { return false; }

            // the fun begins: we need to infer as much type information as we can from the RHS expression.

            var expression = letStmtContext.expression();
            if (expression == null) { Debug.Assert(false, "RHS expression is empty? What's going on here?"); }

            // If RHS is New-ing up an object instance, Set keyword is required & missing.
            if (expression is VBAParser.NewExprContext) { return true; }

            // If RHS is assigning to Nothing (i.e. an "object literal" identifier), Set keyword is required & missing.
            var literalExpression = expression as VBAParser.LiteralExprContext;
            if (literalExpression?.literalExpression()?.literalIdentifier()?.objectLiteralIdentifier() != null) { return true; }

            // note: rhsRefs is correct, but the type of rhsRefs.LastOrDefault may not *necessarily* be the type of the expression.
            // todo: try to *actually* resolve the RHS expression.
            var rhsRefs = state.DeclarationFinder.IdentifierReferences(reference.ParentScoping.QualifiedName)
                .Where(r => !Equals(r, reference) && r.Context.GetAncestor<VBAParser.LetStmtContext>() == letStmtContext);
            var lastRef = rhsRefs.LastOrDefault();
            if (lastRef?.Declaration.IsObject ?? false)
            {
                var typeDeclaration = lastRef.Declaration.AsTypeDeclaration as ClassModuleDeclaration;
                return !HasParameterlessDefaultMember(typeDeclaration);
            }

            return false;
        }

        private static bool HasParameterlessDefaultMember(ClassModuleDeclaration declaration)
        {
            if (declaration?.DefaultMember == null)
            {
                return false;
            }
            var parameters = (declaration.DefaultMember as IParameterizedDeclaration)?.Parameters.ToArray() ?? Enumerable.Empty<ParameterDeclaration>().ToArray();
            return !parameters.Any() || parameters.All(p => p.IsOptional);
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
