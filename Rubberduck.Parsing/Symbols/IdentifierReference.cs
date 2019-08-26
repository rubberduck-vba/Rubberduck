﻿using Antlr4.Runtime;
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
            IEnumerable<IAnnotation> annotations = null,
            bool isSetAssigned = false,
            bool isDefaultMemberAccess = false,
            bool isArrayAccess = false)
        {
            ParentScoping = parentScopingDeclaration;
            ParentNonScoping = parentNonScopingDeclaration;
            QualifiedModuleName = qualifiedName;
            IdentifierName = identifierName;
            Selection = selection;
            Context = context;
            Declaration = declaration;
            HasExplicitLetStatement = hasExplicitLetStatement;
            IsAssignment = isAssignmentTarget;
            IsSetAssignment = isSetAssigned;
            IsDefaultMemberAccess = isDefaultMemberAccess;
            IsArrayAccess = isArrayAccess;
            Annotations = annotations ?? new List<IAnnotation>();
        }

        public QualifiedModuleName QualifiedModuleName { get; }

        public string IdentifierName { get; }

        public Selection Selection { get; }

        /// <summary>
        /// Gets the scoping <see cref="Declaration"/> that contains this identifier reference,
        /// e.g. a module, procedure, function or property.
        /// </summary>
        public Declaration ParentScoping { get; }

        /// <summary>
        /// Gets the non-scoping <see cref="Declaration"/> that contains this identifier reference,
        /// e.g. a user-defined or enum type. Gets the <see cref="ParentScoping"/> if not applicable.
        /// </summary>
        public Declaration ParentNonScoping { get; }

        public bool IsAssignment { get; }

        public bool IsSetAssignment { get; }

        public bool IsDefaultMemberAccess { get; }

        public bool IsArrayAccess { get; }

        public ParserRuleContext Context { get; }

        public Declaration Declaration { get; }

        public IEnumerable<IAnnotation> Annotations { get; }

        public bool HasExplicitLetStatement { get; }

        public bool HasExplicitCallStatement() => Context.Parent is VBAParser.CallStmtContext && ((VBAParser.CallStmtContext)Context).CALL() != null;

        private bool? _hasTypeHint;
        public bool HasTypeHint()
        {
            if (_hasTypeHint.HasValue)
            {
                return _hasTypeHint.Value;
            }

            return HasTypeHint(out _);
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

            var hint = Context.Parent is VBAParser.TypedIdentifierContext typedIdentifierContext
                ? typedIdentifierContext.typeHint()
                : null;

            token = hint?.GetText();
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
