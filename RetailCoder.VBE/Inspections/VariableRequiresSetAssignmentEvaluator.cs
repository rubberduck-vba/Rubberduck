using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections
{
    public static class VariableRequiresSetAssignmentEvaluator
    {
        public static IEnumerable<Declaration> GetDeclarationsPotentiallyRequiringSetAssignment(IEnumerable<Declaration> declarations)
        {
            return declarations.Where(item => MayRequireAssignmentUsingSet(item));
        }

        public static bool RequiresSetAssignment(IdentifierReference reference, IEnumerable<Declaration> declarations)
        {
            var mayRequireAssignmentUsingSet = MayRequireAssignmentUsingSet(reference.Declaration);

            if(!mayRequireAssignmentUsingSet) { return false; }

            var allInterestingDeclarations = GetDeclarationsPotentiallyRequiringSetAssignment(declarations);

            return ObjectOrVariantRequiresSetAssignment(reference, allInterestingDeclarations);
        }

        private static bool MayRequireAssignmentUsingSet(Declaration declaration)
        {
            //The SymbolList includes Variant - which may require 'Set'
            if(SymbolList.ValueTypes.Contains(declaration.AsTypeName))
            {
                return declaration.AsTypeName == Tokens.Variant;
            }

            return 
                TypeIsLikelyAnObject(declaration)
                 && ((IsVariableOrParameter(declaration)
                        && !declaration.IsSelfAssigned)
                || (IsMemberWithReturnType(declaration)
                        && declaration.IsTypeSpecified));
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

        private static bool TypeIsLikelyAnObject(Declaration item)
        {
            var result = !item.IsArray
                         && TypeRequiresSetAssignment(item);
            return result;
        }

        private static bool TypeRequiresSetAssignment(Declaration item)
        {
            if(item.AsTypeDeclaration != null)
            {
                var result = !(ClassModuleDeclaration.HasDefaultMember(item.AsTypeDeclaration)
                    || item.AsTypeDeclaration.DeclarationType == DeclarationType.Enumeration)
                    || item.AsTypeDeclaration.DeclarationType == DeclarationType.UserDefinedType;
                return result;
            }
            return true;    //unit tests: AsTypeDeclaration is often null
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
