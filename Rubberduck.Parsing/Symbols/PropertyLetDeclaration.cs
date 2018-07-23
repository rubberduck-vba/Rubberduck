﻿using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class PropertyLetDeclaration : Declaration, IParameterizedDeclaration, ICanBeDefaultMember
    {
        private readonly List<ParameterDeclaration> _parameters;

        public PropertyLetDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            Accessibility accessibility,
            ParserRuleContext context,
            Selection selection,
            bool isUserDefined,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(
                  name,
                  parent,
                  parentScope,
                  asTypeName,
                  null,
                  false,
                  false,
                  accessibility,
                  DeclarationType.PropertyLet,
                  context,
                  selection,
                  false,
                  null,
                  isUserDefined,
                  annotations,
                  attributes)
        {
            _parameters = new List<ParameterDeclaration>();
        }

        public PropertyLetDeclaration(ComMember member, Declaration parent, QualifiedModuleName module,
            Attributes attributes)
            : this(
                module.QualifyMemberName(member.Name),
                parent,
                parent,
                string.Empty, //TODO:  Need to get the types for these.
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                null,
                attributes)
        {
            _parameters =
                member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module))
                    .ToList(); 
        }

        public PropertyLetDeclaration(ComField field, Declaration parent, QualifiedModuleName module,
            Attributes attributes)
            : this(
                module.QualifyMemberName(field.Name),
                parent,
                parent,
                field.TypeName,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                null,
                attributes)
        { }

        public IEnumerable<ParameterDeclaration> Parameters => _parameters.ToList();

        public void AddParameter(ParameterDeclaration parameter)
        {
            _parameters.Add(parameter);
        }

        /// <summary>
        /// Gets an attribute value indicating whether a member is a class' default member.
        /// If this value is true, any reference to an instance of the class it's the default member of,
        /// should count as a member call to this member.
        /// </summary>
        public bool IsDefaultMember => Attributes.Any(a => a.Name == $"{IdentifierName}.VB_UserMemId" && a.Values.Single() == "0");

        public override bool IsObject => 
            base.IsObject || (Parameters.OrderBy(p => p.Selection).LastOrDefault()?.IsObject ?? false);
    }
}
