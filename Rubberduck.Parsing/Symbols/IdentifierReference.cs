using Antlr4.Runtime;
using Microsoft.CSharp.RuntimeBinder;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReference
    {
        public IdentifierReference(QualifiedModuleName qualifiedName, Declaration parentScopingDeclaration, Declaration parentNonScopingDeclaration, string identifierName,
            Selection selection, ParserRuleContext context, Declaration declaration, bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false, string annotations = null)
        {
            _parentScopingDeclaration = parentScopingDeclaration;
            _parentNonScopingDeclaration = parentNonScopingDeclaration;
            _qualifiedName = qualifiedName;
            _identifierName = identifierName;
            _selection = selection;
            _context = context;
            _declaration = declaration;
            _hasExplicitLetStatement = hasExplicitLetStatement;
            _isAssignmentTarget = isAssignmentTarget;
            _annotations = annotations ?? string.Empty;
        }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedModuleName { get { return _qualifiedName; } }

        private readonly string _identifierName;
        public string IdentifierName { get { return _identifierName; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }

        private readonly Declaration _parentScopingDeclaration;
        /// <summary>
        /// Gets the scoping <see cref="Declaration"/> that contains this identifier reference,
        /// e.g. a module, procedure, function or property.
        /// </summary>
        public Declaration ParentScoping { get { return _parentScopingDeclaration; } }

        private readonly Declaration _parentNonScopingDeclaration;
        /// <summary>
        /// Gets the non-scoping <see cref="Declaration"/> that contains this identifier reference,
        /// e.g. a user-defined or enum type. Gets the <see cref="ParentScoping"/> if not applicable.
        /// </summary>
        public Declaration ParentNonScoping { get { return _parentNonScopingDeclaration; } }

        private readonly bool _isAssignmentTarget;
        public bool IsAssignment { get { return _isAssignmentTarget; } }

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }

        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }

        private readonly string _annotations;
        public string Annotations { get { return _annotations ?? string.Empty; } }

        public bool IsInspectionDisabled(string inspectionName)
        {
            return Annotations.Contains(Grammar.Annotations.IgnoreInspection)
                && Annotations.Contains(inspectionName);
        }

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

        private bool? _hasTypeHint;
        public bool HasTypeHint()
        {
            if (_hasTypeHint.HasValue)
            {
                return _hasTypeHint.Value;
            }

            string token;
            return HasTypeHint(out token);
        }

        public bool HasTypeHint(out string token)
        {
            if (Context == null)
            {
                token = null;
                _hasTypeHint = false;
                return false;
            }

            var method = Context.Parent.GetType().GetMethod("typeHint");
            if (method == null)
            {
                token = null;
                _hasTypeHint = false;
                return false;
            }

            var hint = ((dynamic)Context.Parent).typeHint() as VBAParser.TypeHintContext;
            token = hint == null ? null : hint.GetText();
            _hasTypeHint = hint != null;
            return _hasTypeHint.Value;
        }

        public bool IsSelected(QualifiedSelection selection)
        {
            return QualifiedModuleName == selection.QualifiedName &&
                   Selection.ContainsFirstCharacter(selection.Selection);
        }
    }
}