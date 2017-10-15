﻿using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class PropertyGetDeclaration : Declaration, IParameterizedDeclaration, ICanBeDefaultMember
    {
        private readonly List<Declaration> _parameters;

        public PropertyGetDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            Accessibility accessibility,
            ParserRuleContext context,
            Selection selection,
            bool isArray,
            bool isUserDefined,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(
                  name,
                  parent,
                  parentScope,
                  asTypeName,
                  typeHint,
                  false,
                  false,
                  accessibility,
                  DeclarationType.PropertyGet,
                  context,
                  selection,
                  isArray,
                  asTypeContext,
                  isUserDefined,
                  annotations,
                  attributes)
        {
            _parameters = new List<Declaration>();
        }

        public PropertyGetDeclaration(ComMember member, Declaration parent, QualifiedModuleName module,
            Attributes attributes)
            : this(
                module.QualifyMemberName(member.Name),
                parent,
                parent,
                member.ReturnType.TypeName,
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                member.ReturnType.IsArray,
                false,
                null,
                attributes)
        {
            _parameters =
                member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module))
                    .Cast<Declaration>()
                    .ToList();
        } 

        public IEnumerable<Declaration> Parameters => _parameters.ToList();

        public void AddParameter(Declaration parameter)
        {
            _parameters.Add(parameter);
        }

        /// <summary>
        /// Gets an attribute value indicating whether a member is a class' default member.
        /// If this value is true, any reference to an instance of the class it's the default member of,
        /// should count as a member call to this member.
        /// </summary>
        public bool IsDefaultMember => Attributes.Any(a => a.Name == $"{IdentifierName}.VB_UserMemId" && a.Values.Single() == "0");
    }
}
