using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public abstract class ModuleBodyElementDeclaration : Declaration, IParameterizedDeclaration, ICanBeInterfaceMember, ICanBeDefaultMember
    {
        protected ModuleBodyElementDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            Accessibility accessibility,
            DeclarationType type,
            ParserRuleContext context,
            ParserRuleContext attributesPassContext,
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
                type,
                context,
                attributesPassContext,
                selection,
                isArray,
                asTypeContext,
                isUserDefined,
                annotations,
                attributes)
        {
            if (parent is ClassModuleDeclaration classModule)
            {
                classModule.AddMember(this);
            }
        }
   
        public bool IsInterfaceImplementation => (_implementsResolved ? _implements : InterfaceImplemented) != null;

        private bool _implementsResolved;
        private ClassModuleDeclaration _implements;
        public ClassModuleDeclaration InterfaceImplemented
        {
            get
            {
                if (_implementsResolved)
                {
                    return _implements;
                }

                _implementsResolved = true;
                if (!(ParentDeclaration is ClassModuleDeclaration classModule))
                {
                    return null;
                }

                var identifiers = IdentifierName.Split('_');

                if (identifiers.Length == 1)
                {
                    return null;
                }

                var supertype = string.Join("_", identifiers.Take(identifiers.Length - 1));

                _implements = classModule.Supertypes.Cast<ClassModuleDeclaration>().FirstOrDefault(intrface =>
                    intrface.IdentifierName.Equals(supertype)
                    && intrface.References.Any(reference => ReferenceEquals(reference.ParentScoping, classModule)));

                _implemented = _implements?.Members.FirstOrDefault(member => Implements(member as ICanBeInterfaceMember));

                return _implements;
            }
        }

        internal void InvalidateInterfaceCache()
        {
            _implementsResolved = false;
            _implemented = null;
            _implements = null;
        }

        private Declaration _implemented;
        public Declaration InterfaceMemberImplemented => _implementsResolved || IsInterfaceImplementation ? _implemented : null;

        protected abstract bool Implements(ICanBeInterfaceMember interfaceMember);

        /// <inheritdoc/>
        public bool IsDefaultMember => this.IsDefaultMember();

        private readonly List<ParameterDeclaration> _parameters = new List<ParameterDeclaration>();
        public IEnumerable<ParameterDeclaration> Parameters => _parameters.ToList();

        public void AddParameter(ParameterDeclaration parameter)
        {
            _parameters.Add(parameter);
        }

        protected void AddParameters(IEnumerable<ParameterDeclaration> parameters)
        {
            _parameters.AddRange(parameters);
        }

        /// <inheritdoc/>
        public bool IsInterfaceMember => this.IsInterfaceMember();

        /// <inheritdoc/>
        public ClassModuleDeclaration InterfaceDeclaration => this.InterfaceDeclaration();
    }
}
