﻿using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class PropertySetDeclaration : Declaration, IParameterizedDeclaration, ICanBeDefaultMember
    {
        private readonly List<Declaration> _parameters;

        public PropertySetDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            Accessibility accessibility,
            ParserRuleContext context,
            Selection selection,
            bool isBuiltIn,
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
                  DeclarationType.PropertySet,
                  context,
                  selection,
                  false,
                  null,
                  isBuiltIn,
                  annotations,
                  attributes)
        {
            _parameters = new List<Declaration>();
        }

        public PropertySetDeclaration(ComMember member, Declaration parent, QualifiedModuleName module,
            Attributes attributes) : this(
                module.QualifyMemberName(member.Name),
                parent,
                parent,
                string.Empty, //TODO:  Need to get the types for these.
                Accessibility.Global,
                null,
                Selection.Home,
                true,
                null,
                attributes)
        {
            _parameters =
                member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module))
                    .Cast<Declaration>()
                    .ToList();
        }

        public IEnumerable<Declaration> Parameters
        {
            get
            {
                return _parameters.ToList();
            }
        }

        public void AddParameter(Declaration parameter)
        {
            _parameters.Add(parameter);
        }

        /// <summary>
        /// Gets an attribute value indicating whether a member is a class' default member.
        /// If this value is true, any reference to an instance of the class it's the default member of,
        /// should count as a member call to this member.
        /// </summary>
        public bool IsDefaultMember
        {
            get
            {
                IEnumerable<string> value;
                if (Attributes.TryGetValue(IdentifierName + ".VB_UserMemId", out value))
                {
                    return value.Single() == "0";
                }

                return false;
            }
        }
    }
}
