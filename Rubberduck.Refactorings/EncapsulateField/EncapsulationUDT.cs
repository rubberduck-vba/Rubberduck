using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulationUDT
    {
        private readonly string _typeIdentifier = "This_Type";
        private readonly string _fieldName = "this";
        private readonly IIndenter _indenter;

        private List<IEncapsulatedFieldDeclaration> _members;

        public EncapsulationUDT(string typeIdentifier, string fieldName, IIndenter indenter)
        {
            _typeIdentifier = typeIdentifier;
            _fieldName = fieldName;
            _indenter = indenter;
            _members = new List<IEncapsulatedFieldDeclaration>();
        }

        public EncapsulationUDT(IIndenter indenter)
        {
            _indenter = indenter;
            _members = new List<IEncapsulatedFieldDeclaration>();
        }

        public void AddMember(IEncapsulatedFieldDeclaration field)
        {
            _members.Add(field);
        }

        public string FieldDeclaration
            => $"Private {_fieldName} As {_typeIdentifier}";

        public string TypeDeclarationBlock
        {
            get
            {
                var members = new List<string>();
                members.Add($"Private Type {_typeIdentifier}");

                foreach (var member in _members)
                {
                    var declaration = $"{Capitalize(member.PropertyName)} As {member.Declaration.AsTypeName}";
                    members.Add(declaration);
                }

                members.Add("End Type");

                return  string.Join(Environment.NewLine, _indenter.Indent(members, true));
            }
        }

        private string Capitalize(string input)
        {
            return $"{char.ToUpperInvariant(input[0]) + input.Substring(1, input.Length - 1)}";
        }
    }
}
