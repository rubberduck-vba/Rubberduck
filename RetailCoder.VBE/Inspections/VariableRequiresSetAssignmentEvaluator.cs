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
            if (!reference.IsAssignment) { return false; }
            
            //We know for sure it DOES NOT use 'Set'
            if (!MayRequireAssignmentUsingSet(reference.Declaration)) { return false; }

            //We know for sure that it DOES use 'Set'
            if (RequiresAssignmentUsingSet(reference.Declaration)) { return true; }

            //We need to look everything to understand the RHS - the assigned reference is probably a Variant 
            var allInterestingDeclarations = GetDeclarationsPotentiallyRequiringSetAssignment(state.AllUserDeclarations);

            return ObjectOrVariantRequiresSetAssignment(reference, allInterestingDeclarations);
        }

        private static bool MayRequireAssignmentUsingSet(Declaration declaration)
        {
            if (declaration.AsTypeName == Tokens.Variant) { return true; }

            if (declaration.IsArray) { return false; }

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
                    && (((IsVariableOrParameter(declaration) && !declaration.IsSelfAssigned)
                        || (IsMemberWithReturnType(declaration)  && declaration.IsTypeSpecified)));
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
            var letStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(objectOrVariantRef.Context);

            //definitely needs to use "Set".  e.g., 'Variant myVar = new Collection'
            if (RHSUsesNew(letStmtContext)) { return true; }

            //If the RHS is the identifierName of one of the 'interesting' declarations, we need to use 'Set'
            //unless the 'interesting' declaration is also a Variant
            var rhsIdentifier = GetRHSIdentifier(letStmtContext);
            return variantAndObjectDeclarations
                   .Where(dec => dec.IdentifierName == rhsIdentifier && dec.AsTypeName != Tokens.Variant).Any();
        }

        private static bool IsLetAssignment(IdentifierReference reference)
        {
            var letStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(reference.Context);
            return (reference.IsAssignment && letStmtContext != null);
        }

        private static bool IsAlreadyAssignedUsingSet(IdentifierReference reference)
        {
            var setStmtContext = ParserRuleContextHelper.GetParent<VBAParser.SetStmtContext>(reference.Context);
            return (reference.IsAssignment && setStmtContext != null && setStmtContext.SET() != null);
        }

        private static string GetRHSIdentifier(VBAParser.LetStmtContext letStmtContext)
        {
            for (var idx = 0; idx < letStmtContext.ChildCount; idx++)
            {
                var child = letStmtContext.GetChild(idx);
                if ((child is VBAParser.LiteralExprContext)
                    || (child is VBAParser.LExprContext))
                {
                    return child.GetText();
                }
            }
            return string.Empty;
        }

        private static bool RHSUsesNew(VBAParser.LetStmtContext letStmtContext)
        {
            for (var idx = 0; idx < letStmtContext.ChildCount; idx++)
            {
                var child = letStmtContext.GetChild(idx);
                if ((child is VBAParser.NewExprContext)
                    || (child is VBAParser.CtNewExprContext))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
