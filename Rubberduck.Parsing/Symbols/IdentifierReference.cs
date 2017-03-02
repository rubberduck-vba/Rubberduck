using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System;

namespace Rubberduck.Parsing.Symbols
{
    [DebuggerDisplay("({IdentifierName}) IsAss:{IsAssignment} | {Selection} ")]
    public class IdentifierReference : IEquatable<IdentifierReference>
    {
        public IdentifierReference(
            QualifiedModuleName qualifiedName, 
            Declaration parentScopingDeclaration, 
            Declaration parentNonScopingDeclaration, 
            string identifierName,
            Selection selection,
            ParserRuleContext context, 
            Declaration declaration, 
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false, 
            IEnumerable<IAnnotation> annotations = null)
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
            _annotations = annotations ?? new List<IAnnotation>();
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

        private readonly IEnumerable<IAnnotation> _annotations;
        public IEnumerable<IAnnotation> Annotations { get { return _annotations; } }

        public bool IsIgnoringInspectionResultFor(string inspectionName)
        {
            var isIgnoredAtModuleLevel =
                Declaration.GetModuleParent(_parentScopingDeclaration).Annotations
                .Any(annotation => annotation.AnnotationType == AnnotationType.IgnoreModule
                    && ((IgnoreModuleAnnotation)annotation).IsIgnored(inspectionName));

            return isIgnoredAtModuleLevel || Annotations.Any(annotation => 
                       annotation.AnnotationType == AnnotationType.Ignore
                       && ((IgnoreAnnotation) annotation).IsIgnored(inspectionName));
        }

        private readonly bool _hasExplicitLetStatement;
        public bool HasExplicitLetStatement { get { return _hasExplicitLetStatement; } }

        public bool HasExplicitCallStatement()
        {
            return Context.Parent is VBAParser.CallStmtContext && ((VBAParser.CallStmtContext)Context).CALL() != null;
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
            var method = Context.Parent.GetType().GetMethods().FirstOrDefault(m => m.Name == "typeHint");
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

        public bool Equals(IdentifierReference other)
        {
            return other != null
                && other.QualifiedModuleName.Equals(QualifiedModuleName)
                && other.Selection.Equals(Selection)
                && other.Declaration.Equals(Declaration);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as IdentifierReference);
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(QualifiedModuleName, Selection, Declaration);
        }
    }
}
