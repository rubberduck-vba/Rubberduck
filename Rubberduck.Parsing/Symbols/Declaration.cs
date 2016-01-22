using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines a declared identifier.
    /// </summary>
    [DebuggerDisplay("({DeclarationType}) {Accessibility} {IdentifierName} As {AsTypeName} | {Selection}")]
    public class Declaration : IEquatable<Declaration>
    {
        public Declaration(QualifiedMemberName qualifiedName, Declaration parentDeclaration, string parentScope,
            string asTypeName, bool isSelfAssigned, bool isWithEvents,
            Accessibility accessibility, DeclarationType declarationType, bool isBuiltIn = true, string annotations = null)
            :this(qualifiedName, parentDeclaration, parentScope, asTypeName, isSelfAssigned, isWithEvents, accessibility, declarationType, null, Selection.Home, isBuiltIn, annotations)
        {}

        public Declaration(QualifiedMemberName qualifiedName, Declaration parentDeclaration, string parentScope,
            string asTypeName, bool isSelfAssigned, bool isWithEvents,
            Accessibility accessibility, DeclarationType declarationType, ParserRuleContext context, Selection selection, bool isBuiltIn = false, string annotations = null)
        {
            _qualifiedName = qualifiedName;
            _parentDeclaration = parentDeclaration;
            _parentScope = parentScope;
            _identifierName = qualifiedName.MemberName;
            _asTypeName = asTypeName;
            _isSelfAssigned = isSelfAssigned;
            _isWithEvents = isWithEvents;
            _accessibility = accessibility;
            _declarationType = declarationType;
            _selection = selection;
            _context = context;
            _isBuiltIn = isBuiltIn;
            _annotations = annotations;

            _projectName = _qualifiedName.QualifiedModuleName.ProjectName;

            var ns = Annotations.Split('\n')
                .FirstOrDefault(annotation => annotation.StartsWith(Grammar.Annotations.AnnotationMarker + Grammar.Annotations.Folder));

            string result;
            if (string.IsNullOrEmpty(ns))
            {
                result = _projectName;
            }
            else
            {
                var value = ns.Split(' ')[1];
                result = value;
            }
            _customFolder = result;

            _isArray = IsArray();
            _hasTypeHint = HasTypeHint();
            _isTypeSpecified = IsTypeSpecified();
        }

        /// <summary>
        /// Copies this declaration, optionally modifying one or more values.
        /// </summary>
        /// <returns></returns>
        public Declaration Copy(
            QualifiedMemberName? qualifiedName = null, 
            Declaration parentDeclaration = null, 
            string parentScope = null,
            string asTypeName = null, 
            bool? isSelfAssigned = null, 
            bool? isWithEvents = null,
            Accessibility? accessibility = null, 
            ParserRuleContext context = null, 
            Selection? selection = null, 
            string annotations = null)
        {
            var newQualifiedName = qualifiedName ?? _qualifiedName;
            var newParentDeclaration = parentDeclaration ?? _parentDeclaration;
            var newParentScope = parentScope ?? _parentScope;
            var newAsTypeName = asTypeName ?? _asTypeName;
            var newIsSelfAssigned = isSelfAssigned ?? _isSelfAssigned;
            var newIsWithEvents = isWithEvents ?? _isWithEvents;
            var newAccessibility = accessibility ?? _accessibility;
            var newContext = context ?? _context;
            var newSelection = selection ?? _selection;
            var newAnnotations = annotations ?? _annotations;

            return new Declaration(newQualifiedName, newParentDeclaration, newParentScope, newAsTypeName, newIsSelfAssigned, newIsWithEvents, newAccessibility, _declarationType, newContext, newSelection, _isBuiltIn, newAnnotations);
        }

        /// <summary>
        /// Marks the declaration for the <see cref="IdentifierReferenceResolver"/> to process.
        /// </summary>
        public bool IsDirty { get; set; }

        private readonly bool _isBuiltIn;
        /// <summary>
        /// Marks a declaration as non-user code, e.g. the <see cref="VbaStandardLib"/> or <see cref="ExcelObjectModel"/>.
        /// </summary>
        public bool IsBuiltIn { get { return _isBuiltIn; } }

        private readonly Declaration _parentDeclaration;
        public Declaration ParentDeclaration { get { return _parentDeclaration; } }

        private readonly QualifiedMemberName _qualifiedName;
        public QualifiedMemberName QualifiedName { get { return _qualifiedName; } }

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }

        private ConcurrentBag<IdentifierReference> _references = new ConcurrentBag<IdentifierReference>();
        public IEnumerable<IdentifierReference> References { get { return _references.ToList(); } }

        private readonly ConcurrentBag<IdentifierReference> _memberCalls = new ConcurrentBag<IdentifierReference>();
        public IEnumerable<IdentifierReference> MemberCalls { get { return _memberCalls.ToList(); } }

        public void ClearReferences()
        {
            _references = new ConcurrentBag<IdentifierReference>();
        }

        private readonly string _annotations;
        public string Annotations { get { return _annotations ?? string.Empty; } }

        public bool IsInspectionDisabled(string inspectionName)
        {
            return Annotations.Contains(Grammar.Annotations.IgnoreInspection) 
                && Annotations.Contains(inspectionName);
        }

        public void AddReference(IdentifierReference reference)
        {
            if (reference == null || reference.Declaration.Context == reference.Context)
            {
                return;
            }

            if (reference.Context.Parent != _context 
                && !_references.Select(r => r.Context).Contains(reference.Context.Parent)
                && !_references.Any(r => r.QualifiedModuleName == reference.QualifiedModuleName 
                    && r.Selection.StartLine == reference.Selection.StartLine
                    && r.Selection.EndLine == reference.Selection.EndLine
                    && r.Selection.StartColumn == reference.Selection.StartColumn
                    && r.Selection.EndColumn == reference.Selection.EndColumn))
            {
                _references.Add(reference);
            }
        }

        public void AddMemberCall(IdentifierReference reference)
        {
            if (reference == null || reference.Declaration == null || reference.Declaration.Context == reference.Context)
            {
                return;
            }

            _memberCalls.Add(reference);
        }

        private readonly Selection _selection;
        /// <summary>
        /// Gets a <c>Selection</c> representing the position of the declaration in the code module.
        /// </summary>
        /// <remarks>
        /// Returns <c>default(Selection)</c> for module identifiers.
        /// </remarks>
        public Selection Selection { get { return _selection; } }

        public QualifiedSelection QualifiedSelection { get { return new QualifiedSelection(_qualifiedName.QualifiedModuleName, _selection); } }

        /// <summary>
        /// Gets a reference to the VBProject the declaration is made in.
        /// </summary>
        /// <remarks>
        /// This property is intended to differenciate identically-named VBProjects.
        /// </remarks>
        public VBProject Project { get { return _qualifiedName.QualifiedModuleName.Project; } }

        private readonly string _projectName;
        /// <summary>
        /// Gets the name of the VBProject the declaration is made in.
        /// </summary>
        public string ProjectName { get { return _projectName; } }

        /// <summary>
        /// Gets the name of the VBComponent the declaration is made in.
        /// </summary>
        public string ComponentName { get { return _qualifiedName.QualifiedModuleName.ComponentName; } }

        private readonly string _parentScope;
        /// <summary>
        /// Gets the parent scope of the declaration.
        /// </summary>
        public string ParentScope { get { return _parentScope; } }

        private readonly string _identifierName;
        /// <summary>
        /// Gets the declared name of the identifier.
        /// </summary>
        public string IdentifierName { get { return _identifierName; } }

        private readonly string _asTypeName;
        /// <summary>
        /// Gets the name of the declared type.
        /// </summary>
        /// <remarks>
        /// This value is <c>null</c> if not applicable, 
        /// and <c>Variant</c> if applicable but unspecified.
        /// </remarks>
        public string AsTypeName { get { return _asTypeName; } }

        private bool? _isArray;

        public bool IsArray()
        {
            if (Context == null || _isArray.HasValue && !_isArray.Value)
            {
                return false;
            }

            if (_isArray.HasValue)
            {
                return _isArray.Value;
            }

            //var variableContext = Context as VBAParser.VariableSubStmtContext;
            //if (variableContext != null)
            //{
            //    return variableContext.LPAREN() != null && variableContext.RPAREN() != null;
            //}

            //var typeElementContext = Context as VBAParser.TypeStmt_ElementContext;
            //if (typeElementContext != null)
            //{
            //    return typeElementContext.LPAREN() != null && typeElementContext.RPAREN() != null;
            //}
            //return false;

            try
            {
                var declaration = (dynamic)Context;
                return declaration.LPAREN() != null && declaration.RPAREN() != null;
            }
            catch (RuntimeBinderException)
            {
                return false;
            }
        }

        private bool? _isTypeSpecified;

        public bool IsTypeSpecified()
        {
            if (Context == null || _isTypeSpecified.HasValue && !_isTypeSpecified.Value)
            {
                return false;
            }

            if (_isTypeSpecified.HasValue)
            {
                return _isTypeSpecified.Value;
            }

            //var variableContext = Context as VBAParser.VariableSubStmtContext;
            //if (variableContext != null)
            //{
            //    return variableContext.asTypeClause() != null || HasTypeHint();
            //}

            //var argContext = Context as VBAParser.ArgContext;
            //if (argContext != null)
            //{
            //    return argContext.asTypeClause() != null || HasTypeHint();
            //}

            //var constContext = Context as VBAParser.ConstSubStmtContext;
            //if (constContext != null)
            //{
            //    return constContext.asTypeClause() != null || HasTypeHint();
            //}

            //var functionContext = Context as VBAParser.FunctionStmtContext;
            //if (functionContext != null)
            //{
            //    return functionContext.asTypeClause() != null || HasTypeHint();
            //}

            //var getterContext = Context as VBAParser.PropertyGetStmtContext;
            //if (getterContext != null)
            //{
            //    return getterContext.asTypeClause() != null || HasTypeHint();
            //}

            //var typeElementContext = Context as VBAParser.TypeStmt_ElementContext;
            //if (typeElementContext != null)
            //{
            //    return typeElementContext.asTypeClause() != null || HasTypeHint();
            //}

            //return false;

            try
            {
                var asType = ((dynamic)Context).asTypeClause() as VBAParser.AsTypeClauseContext;
                return asType != null || HasTypeHint();
            }
            catch (RuntimeBinderException)
            {
                return false;
            }
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
                return false;
            }

            //VBAParser.TypeHintContext hint = null;
            //var variableContext = Context as VBAParser.VariableSubStmtContext;
            //if (variableContext != null)
            //{
            //    hint = variableContext.typeHint();
            //}

            ////var argContext = Context as VBAParser.ArgContext;
            ////if (argContext != null)
            ////{
            ////    hint = argContext.typeHint();
            ////}

            //var constContext = Context as VBAParser.ConstSubStmtContext;
            //if (constContext != null)
            //{
            //    hint = constContext.typeHint();
            //}

            //var functionContext = Context as VBAParser.FunctionStmtContext;
            //if (functionContext != null)
            //{
            //    hint = functionContext.typeHint();
            //}

            //var getterContext = Context as VBAParser.PropertyGetStmtContext;
            //if (getterContext != null)
            //{
            //    hint = getterContext.typeHint();
            //}

            ////var typeElementContext = Context as VBAParser.TypeStmt_ElementContext;
            ////if (typeElementContext != null)
            ////{
            ////    hint = typeElementContext.typeHint();
            ////}

            //token = hint == null ? null : hint.GetText();
            //return hint != null;


            try
            {
                var hint = ((dynamic)Context).typeHint() as VBAParser.TypeHintContext;
                token = hint == null ? null : hint.GetText();
                return hint != null;
            }
            catch (RuntimeBinderException)
            {
                token = null;
                return false;
            }
        }

        public bool IsSelected(QualifiedSelection selection)
        {
            return QualifiedName.QualifiedModuleName == selection.QualifiedName &&
                   Selection.ContainsFirstCharacter(selection.Selection);
        }

        private readonly bool _isSelfAssigned;
        /// <summary>
        /// Gets a value indicating whether the declaration is a joined assignment (e.g. "As New xxxxx")
        /// </summary>
        public bool IsSelfAssigned { get { return _isSelfAssigned; } }

        private readonly Accessibility _accessibility;
        /// <summary>
        /// Gets a value specifying the declaration's visibility.
        /// This value is used in determining the declaration's scope.
        /// </summary>
        public Accessibility Accessibility { get { return _accessibility; } }

        private readonly DeclarationType _declarationType;
        /// <summary>
        /// Gets a value specifying the type of declaration.
        /// </summary>
        public DeclarationType DeclarationType { get { return _declarationType; } }

        private readonly bool _isWithEvents;
        /// <summary>
        /// Gets a value specifying whether the declared type is an event provider.
        /// </summary>
        /// <remarks>
        /// WithEvents declarations are used to identify event handler procedures in a module.
        /// </remarks>
        public bool IsWithEvents { get { return _isWithEvents; } }

        /// <summary>
        /// Returns a string representing the scope of an identifier.
        /// </summary>
        public string Scope
        {
            get
            {
                switch (_declarationType)
                {
                    case DeclarationType.Project:
                        return "VBE";
                    case DeclarationType.Class:
                    case DeclarationType.Module:
                        return _qualifiedName.QualifiedModuleName.ToString();
                    case DeclarationType.Procedure:
                    case DeclarationType.Function:
                    case DeclarationType.PropertyGet:
                    case DeclarationType.PropertyLet:
                    case DeclarationType.PropertySet:
                        return _qualifiedName.QualifiedModuleName + "." + _identifierName;
                    case DeclarationType.Event:
                        return _parentScope + "." + _identifierName;
                    default:
                        return _parentScope;
                }
            }
        }

        private readonly string _customFolder;
        public string CustomFolder
        {
            get
            {
                return _customFolder;
            }
        }

        public bool Equals(Declaration other)
        {
            return other.Project == Project
                && other.IdentifierName == IdentifierName
                && other.DeclarationType == DeclarationType
                && other.Scope == Scope
                && other.ParentScope == ParentScope;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as Declaration);
        }

        public override int GetHashCode()
        {
            return string.Concat(QualifiedName.QualifiedModuleName.ProjectHashCode, _identifierName, _declarationType, Scope, _parentScope).GetHashCode();
        }
    }
}
