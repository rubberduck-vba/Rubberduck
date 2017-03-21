using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections
{
    class VariableRequiresSetAssignmentEvaluator
    {
        private readonly RubberduckParserState _parserState;
        public VariableRequiresSetAssignmentEvaluator(RubberduckParserState parserState)
        {
            _parserState = parserState;
        }

        public IEnumerable<Declaration> GetDeclarationsPotentiallyRequiringSetAssignment()
        {
            var interestingDeclarations = _parserState.AllUserDeclarations.Where(item =>
                    IsVariableOrParameter(item)
                    && !item.IsSelfAssigned
                    && TypeIsAnObjectOrVariant(item));

            var interestingMembers = _parserState.AllUserDeclarations.Where(item =>
                    IsMemberWithReturnType(item)
                    && item.IsTypeSpecified
                    && TypeIsAnObjectOrVariant(item));

            var allInterestingDeclarations = interestingDeclarations
                    .Union(HasReturnAssignment(interestingMembers));

            return allInterestingDeclarations;
        }

        public bool RequiresSetAssignment(IdentifierReference reference)
        {
            var declaration = reference.Declaration;
            var MayRequireAssignmentUsingSet =
                 (IsVariableOrParameter(declaration) || IsMemberWithReturnType(declaration) )
                 && !declaration.IsSelfAssigned
                 && TypeIsAnObjectOrVariant(declaration);

            if(!MayRequireAssignmentUsingSet) { return false; }

            var allInterestingDeclarations = GetDeclarationsPotentiallyRequiringSetAssignment();

            return ObjectOrVariantRequiresSetAssignment(reference, allInterestingDeclarations);
        }

        private bool IsMemberWithReturnType(Declaration item)
        {
            return (item.DeclarationType == DeclarationType.Function
                || item.DeclarationType == DeclarationType.PropertyGet);
        }

        private IEnumerable<Declaration> HasReturnAssignment(IEnumerable<Declaration> interestingMembers)
        {
            return interestingMembers.SelectMany(member =>
                      member.References.Where(memberRef => memberRef.ParentScoping.Equals(member)
                           && memberRef.IsAssignment)).Select(reference => reference.Declaration);
        }

        private bool IsVariableOrParameter(Declaration item)
        {
            return item.DeclarationType == DeclarationType.Variable
                    || item.DeclarationType == DeclarationType.Parameter;
        }

        private bool TypeIsAnObjectOrVariant(Declaration item)
        {
            return !item.IsArray
                    && !ValueOnlyTypes().Contains(item.AsTypeName)
                    && (item.AsTypeDeclaration == null
                        || TypeRequiresSetAssignment(item));
        }

        private IEnumerable<string> ValueOnlyTypes()
        {
            var nonSetTypes = SymbolList.ValueTypes.ToList();
            nonSetTypes.Remove(Tokens.Variant);
            return nonSetTypes;
        }

        private bool TypeRequiresSetAssignment(Declaration item)
        {
            return (!ClassModuleDeclaration.HasDefaultMember(item.AsTypeDeclaration))
                && (item.AsTypeDeclaration.DeclarationType != DeclarationType.Enumeration
                && item.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType
                    && item.AsTypeDeclaration != null);
        }

        private bool ObjectOrVariantRequiresSetAssignment(IdentifierReference variantOrObjectRef, IEnumerable<Declaration> variantAndObjectDeclarations)
        {
            //Not an assignment...not interested
            if (!variantOrObjectRef.IsAssignment)
            {
                return false;
            }

            //Already assigned using 'Set'
            if (IsSetAssignment(variantOrObjectRef)) { return true; };

            if (variantOrObjectRef.Declaration.AsTypeName != Tokens.Variant) { return true; }
            
            //Variants can be assigned with or without 'Set' depending...
            var letStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(variantOrObjectRef.Context);

            //definitely needs to use "Set".  e.g., 'Variant myVar = new Collection'
            if (RHSUsesNew(letStmtContext)) { return true; }

            //If the RHS is the identifierName of one of the 'interesting' declarations, we need to use 'Set'
            //unless the 'interesting' declaration is also a Variant
            var rhsIdentifier = GetRHSIdentifier(letStmtContext);
            return variantAndObjectDeclarations
                   .Where(dec => dec.IdentifierName == rhsIdentifier && dec.AsTypeName != Tokens.Variant).Any();
        }


        private bool IsLetAssignment(IdentifierReference reference)
        {
            var letStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(reference.Context);
            return (reference.IsAssignment && letStmtContext != null);
        }

        private bool IsSetAssignment(IdentifierReference reference)
        {
            var setStmtContext = ParserRuleContextHelper.GetParent<VBAParser.SetStmtContext>(reference.Context);
            return (reference.IsAssignment && setStmtContext != null && setStmtContext.SET() != null);
        }

        private string GetRHSIdentifier(VBAParser.LetStmtContext letStmtContext)
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

        private bool RHSUsesNew(VBAParser.LetStmtContext letStmtContext)
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
