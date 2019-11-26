using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class UDTGenerator
    {
        private readonly string _typeIdentifier = "This_Type";
        private readonly IIndenter _indenter;

        private List<IEncapsulatedFieldDeclaration> _members;

        public UDTGenerator(string typeIdentifier, IIndenter indenter)
            :this(indenter)
        {
            _typeIdentifier = typeIdentifier;
        }

        public UDTGenerator(IIndenter indenter)
        {
            _indenter = indenter;
            _members = new List<IEncapsulatedFieldDeclaration>();
        }

        public string AsTypeName => _typeIdentifier;

        public void AddMember(IEncapsulatedFieldDeclaration field)
        {
            _members.Add(field);
        }

        public string FieldDeclaration(string identifier, Accessibility accessibility = Accessibility.Private)
            => $"{accessibility} {identifier} {Tokens.As} {_typeIdentifier}";

        public string TypeDeclarationBlock
        {
            get
            {
                var members = new List<string>();
                members.Add($"{Tokens.Private} {Tokens.Type} {_typeIdentifier}");

                foreach (var member in _members)
                {
                    var declaration = $"{member.PropertyName.Capitalize()} As {member.AsTypeName}";
                    members.Add(declaration);
                }

                members.Add("End Type");

                return  string.Join(Environment.NewLine, _indenter.Indent(members, true));
            }
        }
    }
}
