using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulationUDT
    {
        private const string _typeIdentifier = "This_Type";
        private const string _fieldName = "this";

        private List<IEncapsulatedFieldDeclaration> _members;

        public EncapsulationUDT()
        {
            _members = new List<IEncapsulatedFieldDeclaration>();
        }

        public void AddMember(IEncapsulatedFieldDeclaration field)
        {
            _members.Add(field);
        }

        public string FieldDeclaration
            => $"Private {_fieldName} As {_typeIdentifier}";

        public string DeclarationAndField
        {
            get
            {
                var members = new List<string>();
                foreach (var member in _members)
                {
                    var declaration = $"{Capitalize(member.PropertyName)} As {member.Declaration.AsTypeName}";
                    members.Add(declaration);
                }
                return
$@"Private Type {_typeIdentifier}
    {string.Join(Environment.NewLine, members)}
End Type

{FieldDeclaration}{Environment.NewLine}";
            }
        }

        private string Capitalize(string input)
        {
            return $"{char.ToUpperInvariant(input[0]) + input.Substring(1, input.Length - 1)}";
        }
    }
}
