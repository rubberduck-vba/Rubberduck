using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReference
    {
        public IdentifierReference(QualifiedModuleName qualifiedName, string identifierName, Selection selection, ParserRuleContext context, Declaration declaration, bool isAssignmentTarget = false)
        {
            _qualifiedName = qualifiedName;
            _identifierName = identifierName;
            _selection = selection;
            _context = context;
            _declaration = declaration;
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

        public bool HasExplicitLetStatement()
        {
            var context = FindValueAssignmentContext(Context);
            if (context == null)
            {
                return false;
            }

            var statement = context.LET();
            return statement != null && statement.Symbol.Text == Tokens.Let;
        }

        private VBAParser.LetStmtContext FindValueAssignmentContext(RuleContext context)
        {
            var statement = context.Parent as VBAParser.LetStmtContext;
            if (statement != null && context is VBAParser.ImplicitCallStmt_InStmtContext)
            {
                return statement;
            }

            var parent = context.Parent;
            if (parent != null)
            {
                return FindValueAssignmentContext(parent);
            }

            return null;
        }

        private VBAParser.SetStmtContext FindReferenceAssignmentContext(RuleContext context)
        {
            var statement = context.Parent as VBAParser.SetStmtContext;
            if (statement != null && context is VBAParser.ImplicitCallStmt_InStmtContext)
            {
                return statement;
            }

            var parent = context.Parent;
            if (parent != null)
            {
                return FindReferenceAssignmentContext(parent);
            }

            return null;
        }

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
            try
            {
                var hint = ((dynamic) Context.Parent).typeHint();
                return hint != null && !string.IsNullOrEmpty(hint.GetText());
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}