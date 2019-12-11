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
        private readonly string _typeIdentifierName;

        private List<IEncapsulateFieldCandidate> _members;

        public UDTDeclarationGenerator(string typeIdentifierName)
        {
            _typeIdentifierName = typeIdentifierName;
            _members = new List<IEncapsulateFieldCandidate>();
        }

        public void AddMember(IEncapsulateFieldCandidate field) => _members.Add(field);

        public void AddMembers(IEnumerable<IEncapsulateFieldCandidate> fields) => _members.AddRange(fields);

        public string FieldDeclarationBlock(string identifierName, Accessibility accessibility = Accessibility.Private)
            => $"{accessibility} {identifierName} {Tokens.As} {_typeIdentifierName}";

        public string TypeDeclarationBlock(IIndenter indenter = null, Accessibility accessibility = Accessibility.Private)
        {
            if (indenter != null)
            {
                return string.Join(Environment.NewLine, indenter.Indent(BlockLines(accessibility), true));
            }
            return string.Join(Environment.NewLine, BlockLines(accessibility));
        }

        public IEnumerable<string> BlockLines(Accessibility accessibility)
        {
            var blockLines = new List<string>();

            blockLines.Add($"{accessibility.TokenString()} {Tokens.Type} {_typeIdentifierName}");

            _members.ForEach(m => blockLines.Add($"{m.PropertyName} As {m.AsTypeName}"));

            blockLines.Add("End Type");

            return blockLines;
        } 
    }
}
