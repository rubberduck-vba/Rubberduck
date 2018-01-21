using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
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

        public static bool RequiresSetAssignment(IdentifierReference reference, RubberduckParserState state)
        {
            //Not an assignment...definitely does not require a 'Set' assignment
            if (!reference.IsAssignment)
            {
                return false;
            }
            
            //We know for sure it DOES NOT use 'Set'
            if (!MayRequireAssignmentUsingSet(reference.Declaration))
            {
                return false;
            }

            //We know for sure that it DOES use 'Set'
            if (RequiresAssignmentUsingSet(reference.Declaration))
            {
                return true;
            }

            //We need to look everything to understand the RHS - the assigned reference is probably a Variant 
            var allInterestingDeclarations = GetDeclarationsPotentiallyRequiringSetAssignment(state.AllUserDeclarations);

            return ObjectOrVariantRequiresSetAssignment(reference, allInterestingDeclarations);
        }

        private static bool MayRequireAssignmentUsingSet(Declaration declaration)
        {
            if (declaration.DeclarationType == DeclarationType.PropertyLet)
            {
                return false;
            }

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
                return declaration.AsTypeDeclaration.DeclarationType == DeclarationType.UserDefinedType
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

        private static bool ObjectOrVariantRequiresSetAssignment(IdentifierReference objectOrVariantRef, IEnumerable<Declaration> variantAndObjectDeclarations)
        {
            //Not an assignment...nothing to evaluate
            if (!objectOrVariantRef.IsAssignment)
            {
                return false;
            }

            if (IsAlreadyAssignedUsingSet(objectOrVariantRef)
                    || objectOrVariantRef.Declaration.AsTypeName != Tokens.Variant)
            {
                return true;
            }

            //Variants can be assigned with or without 'Set' depending...
            var letStmtContext = objectOrVariantRef.Context.GetAncestor<VBAParser.LetStmtContext>();

            //A potential error is only possible for let statements: rset, lset and other type specific assignments are always let assignments; 
            //assignemts in for each loop statements are do not require the set keyword.
            if (letStmtContext == null)
            {
                return false;
            }

            //You can only new up objects.
            if (RHSUsesNew(letStmtContext)) { return true; }

            if (RHSIsLiteral(letStmtContext))
            {
                if(RHSIsObjectLiteral(letStmtContext))
                {
                    return true;
                }
                //All literals but the object literal potentially do not need a set assignment.
                //We cannot get more information from the RHS and do not want false positives.
                return false;
            }

            //If the RHS is the identifierName of one of the 'interesting' declarations, we need to use 'Set'
            //unless the 'interesting' declaration is also a Variant
            var rhsIdentifier = GetRHSIdentifierExpressionText(letStmtContext);
            return variantAndObjectDeclarations.Any(dec => dec.IdentifierName == rhsIdentifier && dec.AsTypeName != Tokens.Variant);
        }

        private static bool IsLetAssignment(IdentifierReference reference)
        {
            var letStmtContext = reference.Context.GetAncestor<VBAParser.LetStmtContext>();
            return (reference.IsAssignment && letStmtContext != null);
        }

        private static bool IsAlreadyAssignedUsingSet(IdentifierReference reference)
        {
            var setStmtContext = reference.Context.GetAncestor<VBAParser.SetStmtContext>();
            return (reference.IsAssignment && setStmtContext?.SET() != null);
        }

        private static string GetRHSIdentifierExpressionText(VBAParser.LetStmtContext letStmtContext)
        {
            var expression = letStmtContext.expression();
            return expression is VBAParser.LExprContext ? expression.GetText() : string.Empty;
        }

        private static bool RHSUsesNew(VBAParser.LetStmtContext letStmtContext)
        {
            var expression = letStmtContext.expression();
            return (expression is VBAParser.NewExprContext);
        }

        private static bool RHSIsLiteral(VBAParser.LetStmtContext letStmtContext)
        {
            return letStmtContext.expression() is VBAParser.LiteralExprContext;                   
        }

        private static bool RHSIsObjectLiteral(VBAParser.LetStmtContext letStmtContext)
        {
            var rhsAsLiteralExpr = letStmtContext.expression() as VBAParser.LiteralExprContext;
            return rhsAsLiteralExpr?.literalExpression()?.literalIdentifier()?.objectLiteralIdentifier() != null;
        }
    }
}
