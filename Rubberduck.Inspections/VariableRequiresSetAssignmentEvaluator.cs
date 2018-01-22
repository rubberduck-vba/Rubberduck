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
                // assignment already has a Set keyword
                // (but is it misplaced? ...hmmm... beyond the scope of *this* inspection though)
                // if we're only ever assigning to 'Nothing', might as well flag it though
                if (reference.Declaration.References.Where(r => r.IsAssignment).All(r =>
                    r.Context.GetAncestor<VBAParser.SetStmtContext>().expression().GetText() == Tokens.Nothing))
                {
                    return true;
                }
            }

            var letStmtContext = reference.Context.GetAncestor<VBAParser.LetStmtContext>();
            if (letStmtContext == null)
            {
                // we're probably in a For Each loop
                return false;
            }

            var declaration = reference.Declaration;
            if (declaration.IsArray)
            {
                // arrays don't need a Set statement... todo figure out if array items are objects
                return false;
            }

            var isObjectVariable = declaration.IsObject();
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

        private static bool ObjectOrVariantRequiresSetAssignment(IdentifierReference objectOrVariantRef, IEnumerable<Declaration> variantAndObjectDeclarations)
        {
            //Not an assignment...nothing to evaluate
            if (!objectOrVariantRef.IsAssignment)
            {
                return false;
            }

            if (objectOrVariantRef.Declaration.AsTypeName != Tokens.Variant)
            {
                return true;
            }

            //Variants can be assigned with or without 'Set' depending...
            var letStmtContext = objectOrVariantRef.Context.GetAncestor<VBAParser.LetStmtContext>();

            //A potential error is only possible for let statements: rset, lset and other type specific assignments are always let assignments; 
            //assignemts in for each loop statements are do not require the set keyword.
            if(letStmtContext == null)
            {
                return false;
            }

            //You can only new up objects.
            if (RHSUsesNew(letStmtContext))
            {
                return true;
            }

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
            return variantAndObjectDeclarations.Any(dec => dec.IdentifierName == rhsIdentifier 
                && dec.AsTypeName != Tokens.Variant
                && dec.Attributes.HasDefaultMemberAttribute(dec.IdentifierName, out _));
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
