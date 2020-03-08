using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Parsing.Symbols
{
    public abstract class ModuleBodyElementDeclaration : Declaration, IParameterizedDeclaration, IInterfaceExposable, ICanBeDefaultMember
    {
        protected ModuleBodyElementDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            AsTypeClauseContext asTypeContext,
            string typeHint,
            Accessibility accessibility,
            DeclarationType type,
            ParserRuleContext context,
            ParserRuleContext attributesPassContext,
            Selection selection,
            bool isArray,
            bool isUserDefined,
            IEnumerable<IParseTreeAnnotation> annotations,
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
            if (!(parent is ModuleDeclaration module))
            {
                return;
            }

            module.AddMember(this);

            _interface = new Lazy<(bool IsInterfaceImplementation, ClassModuleDeclaration InterfaceImplemented)>(() => ResolveInterface(this), true);
            _implemented = new Lazy<Declaration>(() => MemberImplemented(this), true);
        }

        /// <inheritdoc/>
        public bool IsDefaultMember => this.IsDefaultMember();

        private readonly List<ParameterDeclaration> _parameters = new List<ParameterDeclaration>();
        public IReadOnlyList<ParameterDeclaration> Parameters => _parameters.ToList();

        public void AddParameter(ParameterDeclaration parameter)
        {
            _parameters.Add(parameter);
        }

        protected void AddParameters(IEnumerable<ParameterDeclaration> parameters)
        {
            _parameters.AddRange(parameters);
        }

        private Lazy<(bool IsInterfaceImplementation, ClassModuleDeclaration InterfaceImplemented)> _interface;

        /// <summary>
        /// Returns true if this member is a concrete implementation of an interface.
        /// </summary>
        public bool IsInterfaceImplementation => _interface != null && _interface.Value.IsInterfaceImplementation;

        /// <summary>
        /// Returns the interface that this member implements from, or null if not an implementation.
        /// </summary>
        public ClassModuleDeclaration InterfaceImplemented => _interface?.Value.InterfaceImplemented;

        private Lazy<Declaration> _implemented;

        /// <summary>
        /// Returns the interface member that this member is a concrete implementation of, or null if not an implementation.
        /// </summary>
        public Declaration InterfaceMemberImplemented => _implemented?.Value;

        /// <inheritdoc/>
        public string ImplementingIdentifierName => this.ImplementingIdentifierName();

        /// <summary>
        /// Returns true if this Declaration is a concrete implementation of the passed member.
        /// </summary>
        /// <param name="interfaceMember">The member to test for implementation of.</param>
        /// <returns>True if this Declaration is a concrete implementation of interfaceMember.</returns>
        protected abstract bool Implements(IInterfaceExposable interfaceMember);

        /// <inheritdoc/>
        public bool IsInterfaceMember => this.IsInterfaceMember();

        /// <inheritdoc/>
        public ClassModuleDeclaration InterfaceDeclaration => this.InterfaceDeclaration();

        internal void InvalidateInterfaceCache()
        {
            if (!(ParentDeclaration is ClassModuleDeclaration))
            {
                return;
            }

            _interface = new Lazy<(bool IsInterfaceImplementation, ClassModuleDeclaration InterfaceImplemented)>(() => ResolveInterface(this), true);
            _implemented = new Lazy<Declaration>(() => MemberImplemented(this), true);
        }

        private static (bool IsInterfaceImplementation, ClassModuleDeclaration InterfaceImplemented) ResolveInterface(ModuleBodyElementDeclaration element)
        {
            if (!(element.ParentDeclaration is ClassModuleDeclaration classModule))
            {
                return (false, null);
            }

            var identifiers = element.IdentifierName.Split('_');

            if (identifiers.Length == 1)
            {
                return (false, null);
            }

            /*
             * The following MS-VBAL rule effectively limits the the IdentifierName of an implemented interface member to the last token in 'identifiers':
             *
             * 5.2.4.2 - A class may not be used as an interface class if the names of any of its public variable or method
             * methods contain an underscore character (Unicode u+005F).
             *
             * Note that there is NOT a corresponding restriction on the <class-type-name>. 
             */

            var supertype = string.Join("_", identifiers.Take(identifiers.Length - 1));

            var implements = classModule.Supertypes.Cast<ClassModuleDeclaration>().FirstOrDefault(intrface =>
                intrface.IdentifierName.Equals(supertype)
                && intrface.References.Any(reference => ReferenceEquals(reference.ParentScoping, classModule)));

            if (implements == null)
            {
                return (false, null);
            }

            return implements.Members.Any(member => element.Implements(member as IInterfaceExposable)) ? (true, implements) : (false, null);
        }

        private static Declaration MemberImplemented(ModuleBodyElementDeclaration element)
        {
            return element.InterfaceImplemented?.Members.FirstOrDefault(member => element.Implements(member as IInterfaceExposable));
        }

        public abstract BlockContext Block { get; }
    }
}
