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
    public class UDTDeclarationGenerator
    {
        private readonly string _typeIdentifier = "This_Type";
        private readonly IIndenter _indenter;

        private List<IEncapsulateFieldCandidate> _members;

        public UDTDeclarationGenerator(string typeIdentifier) //, IIndenter indenter)
            : this()
            //:this(indenter)
        {
            _typeIdentifier = typeIdentifier;
        }

        public UDTDeclarationGenerator() //IIndenter indenter)
        {
            //_indenter = indenter;
            _members = new List<IEncapsulateFieldCandidate>();
        }

        public string AsTypeName => _typeIdentifier;

        public void AddMember(IEncapsulateFieldCandidate field)
        {
            _members.Add(field);
        }

        public string FieldDeclaration(string identifier, Accessibility accessibility = Accessibility.Private)
            => $"{accessibility} {identifier} {Tokens.As} {_typeIdentifier}";

        public string TypeDeclarationBlock(IIndenter indenter)
        {
            //get
            //{
                var members = new List<string>();
                members.Add($"{Tokens.Private} {Tokens.Type} {_typeIdentifier}");

                foreach (var member in _members)
                {
                    var declaration = $"{member.PropertyName} As {member.AsTypeName}";
                    members.Add(declaration);
                }

                members.Add("End Type");

                return  string.Join(Environment.NewLine, indenter.Indent(members, true));
            //}
        }
    }
}
