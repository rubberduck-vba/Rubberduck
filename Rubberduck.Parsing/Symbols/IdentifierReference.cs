using System;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.CSharp.RuntimeBinder;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReference
    {
        public IdentifierReference(QualifiedModuleName qualifiedName, string identifierName, Selection selection, ParserRuleContext context, Declaration declaration, bool isAssignmentTarget = false, bool hasExplicitLetStatement = false)
        {
            _qualifiedName = qualifiedName;
            _identifierName = identifierName;
            _selection = selection;
            _context = context;
            _declaration = declaration;
            _hasExplicitLetStatement = hasExplicitLetStatement;
            _isAssignmentTarget = isAssignmentTarget;
        }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedModuleName { get { return _qualifiedName; } }

        private readonly string _identifierName;
        public string IdentifierName { get { return _identifierName; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }

        private readonly bool _isAssignmentTarget;
        public bool IsAssignment { get { return _isAssignmentTarget; } }

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }

        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }

        private readonly bool _hasExplicitLetStatement;
        public bool HasExplicitLetStatement { get { return _hasExplicitLetStatement; } }

        public bool HasExplicitCallStatement()
        {
            var memberProcedureCall = Context.Parent as VBAParser.ECS_MemberProcedureCallContext;
            var procedureCall = Context.Parent as VBAParser.ECS_ProcedureCallContext;

            return HasExplicitCallStatement(memberProcedureCall) || HasExplicitCallStatement(procedureCall);
        }

        private bool HasExplicitCallStatement(VBAParser.ECS_MemberProcedureCallContext call)
        {
            if (call == null)
            {
                return false;
            }
            var statement = call.CALL();
            return statement != null && statement.Symbol.Text == Tokens.Call;
        }

        private bool HasExplicitCallStatement(VBAParser.ECS_ProcedureCallContext call)
        {
            if (call == null)
            {
                return false;
            }
            var statement = call.CALL();
            return statement != null && statement.Symbol.Text == Tokens.Call;
        }

        public bool HasTypeHint()
        {
            string token;
            return HasTypeHint(out token);
        }

        public bool HasTypeHint(out string token)
        {
            if (Context == null)
            {
                token = null;
                return false;
            }

            try
            {
                var hint = ((dynamic)Context.Parent).typeHint() as VBAParser.TypeHintContext;
                token = hint == null ? null : hint.GetText();
                return hint != null;
            }
            catch (RuntimeBinderException)
            {
                token = null;
                return false;
            }
        }
    }
}