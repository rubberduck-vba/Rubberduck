using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences
{
    public class UserDefinedTypeInstance
    {
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

        private List<IdentifierReference> _udtMemberReferences;
        public IReadOnlyCollection<IdentifierReference> UDTMemberReferences => _udtMemberReferences;

        private bool IsRelatedReference(IdentifierReference idRef, IEnumerable<IdentifierReference> fieldReferences)
        {
            if (idRef.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmac))
            {
                var goalContext = wmac.GetAncestor<VBAParser.WithStmtContext>();
                return fieldReferences.Any(rf => HasSameAncestor<VBAParser.WithStmtContext>(rf, goalContext));
            }
            else if (idRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var memberAccessExprContext))
            {
                return fieldReferences.Any(rf => HasSameAncestor<VBAParser.MemberAccessExprContext>(rf, memberAccessExprContext));
            }
            throw new ArgumentOutOfRangeException();
        }

        private bool HasSameAncestor<T>(IdentifierReference idRef, ParserRuleContext goalContext) where T : ParserRuleContext
        {
            Debug.Assert(goalContext != null);
            Debug.Assert(goalContext is VBAParser.MemberAccessExprContext || goalContext is VBAParser.WithStmtContext);

            const int maxGetAncestorAttempts = 100;
            var guard = 0;
            var accessExprContext = idRef.Context.GetAncestor<T>();
            while (accessExprContext != null && ++guard < maxGetAncestorAttempts)
            {
                var prCtxt = accessExprContext as ParserRuleContext;
                if (prCtxt == goalContext)
                {
                    return true;
                }
                accessExprContext = accessExprContext.GetAncestor<T>();
            }

            Debug.Assert(guard < maxGetAncestorAttempts);
            return false;
        }
    }
}
