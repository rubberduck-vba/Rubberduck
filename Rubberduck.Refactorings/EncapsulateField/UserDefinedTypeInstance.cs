using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class UserDefinedTypeInstance
    {
        private readonly List<IdentifierReference> _udtMemberReferences;
        public UserDefinedTypeInstance(VariableDeclaration field, IEnumerable<Declaration> udtMembers)
        {
            if (!(field.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false))
            {
                throw new ArgumentException();
            }

            InstanceField = field;
            _udtMemberReferences = udtMembers.SelectMany(m => m.References)
                .Where(rf => IsRelatedReference(rf, InstanceField.References)).ToList();
        }

        public VariableDeclaration InstanceField { get; }

        public string UserDefinedTypeIdentifier => InstanceField.AsTypeDeclaration.IdentifierName;

        public IReadOnlyCollection<IdentifierReference> UDTMemberReferences => _udtMemberReferences;

        private bool IsRelatedReference(IdentifierReference udtMemberRef, IEnumerable<IdentifierReference> fieldReferences)
        {
            if (udtMemberRef.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmacUdtMemberAncestor))
            {
                if (wmacUdtMemberAncestor.TryGetAncestor<VBAParser.WithStmtContext>(out var wscUdtMemberAncestor))
                {
                    return fieldReferences.Any(fieldRef => AreRelated(fieldRef, wscUdtMemberAncestor));
                }
            }
            else if (udtMemberRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var memberAccessExprContext))
            {
                return fieldReferences.Any(rf => AreRelated(rf, memberAccessExprContext));
            }
            //TODO: These need to be an exception which results in ending the refactoring action
            throw new ArgumentOutOfRangeException();
        }

        private bool AreRelated<T>(IdentifierReference fieldRef, T udtMemberRefCtxt) where T : ParserRuleContext
        {
            //TODO: These need to be an exception -> which results in ending the refactoring action
            Debug.Assert(udtMemberRefCtxt != null);
            Debug.Assert(udtMemberRefCtxt is VBAParser.MemberAccessExprContext || udtMemberRefCtxt is VBAParser.WithStmtContext);

            if (udtMemberRefCtxt is null)
            {
                return false;
            }

            const int maxGetAncestorDescendentAttempts = 100;
            var guard = 0;
            var fieldRefAncestor = fieldRef.Context.GetAncestor<T>();
            while (fieldRefAncestor != null && ++guard < maxGetAncestorDescendentAttempts)
            {
                if (udtMemberRefCtxt is null)
                {
                    return false;
                }
                
                if (fieldRefAncestor is ParserRuleContext prctxt && prctxt.Equals(udtMemberRefCtxt))
                {
                    return true;
                }

                if (fieldRefAncestor is VBAParser.MemberAccessExprContext)
                {
                    udtMemberRefCtxt = udtMemberRefCtxt.GetDescendent<T>();
                    continue;
                }

                udtMemberRefCtxt = udtMemberRefCtxt.GetAncestor<T>();
            }

            //TODO: This needs to be an exception -> which results in ending the refactoring action
            Debug.Assert(guard < maxGetAncestorDescendentAttempts);
            return false;
        }
    }
}
