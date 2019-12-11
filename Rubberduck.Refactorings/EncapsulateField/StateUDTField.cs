using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IStateUDTField
    {
        string TypeIdentifier { set; get; }
        string TypeDeclarationBlock(IIndenter indenter = null);
        string FieldDeclarationBlock { get; }
        void AddMembers(IEnumerable<IEncapsulateFieldCandidate> fields);
        void Reset();
        string NewFieldName { get; }
    }

    public class StateUDTField : EncapsulateFieldCandidate, IUserDefinedTypeCandidate, IStateUDTField
    {
        private const string _defaultTypeIdentifier = "This_Type";
        private const string _defaultNewFieldName = "this";
        private List<IEncapsulateFieldCandidate> _members;

        public StateUDTField(string identifier, string asTypeName, QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
            : base(identifier, asTypeName, qmn, validator)
        {
            PropertyName = identifier;
            NewFieldName = identifier;
            AsTypeName = asTypeName;
            TypeIdentifier = asTypeName;
            _members = new List<IEncapsulateFieldCandidate>();

            PropertyAccessor = AccessorTokens.Field;
            ReferenceAccessor = AccessorTokens.Field;
            //ReferenceQualifier = Parent.ReferenceWithinNewProperty;
        }

        public string TypeIdentifier { set; get; }

        public override string PropertyName { set; get; }

        public override string NewFieldName { set; get; }

        public override bool IsSelfConsistent => _validator.IsValidVBAIdentifier(PropertyName, DeclarationType.Property)
                            && !(PropertyName.EqualsVBAIdentifier(ParameterName)
                                    || PropertyName.EqualsVBAIdentifier(ParameterName));

        public IEnumerable<IUserDefinedTypeMemberCandidate> Members { get; }
        public void AddMember(IUserDefinedTypeMemberCandidate member) { throw new NotImplementedException(); }
        public bool FieldQualifyUDTMemberPropertyName { set; get; }
        public bool TypeDeclarationIsPrivate { set; get; } = true;

        //public void AddMember(IEncapsulateFieldCandidate field) => _members.Add(field);
        public void Reset()
        {
            _members.Clear();
            NewFieldName = _defaultNewFieldName;
            TypeIdentifier = _defaultTypeIdentifier;
        }

        public void AddMembers(IEnumerable<IEncapsulateFieldCandidate> fields) => _members.AddRange(fields);

        public string FieldDeclarationBlock //(string fieldName) // string identifierName, Accessibility accessibility = Accessibility.Private)
            => $"{Accessibility.Private} {NewFieldName} {Tokens.As} {AsTypeName}";

        public string TypeDeclarationBlock(IIndenter indenter = null) //= null, Accessibility accessibility = Accessibility.Private)
        {
            if (indenter != null)
            {
                return string.Join(Environment.NewLine, indenter?.Indent(BlockLines(Accessibility.Private) ?? BlockLines(Accessibility.Private), true));
            }
            return string.Join(Environment.NewLine, BlockLines(Accessibility.Private));
        }

        private IEnumerable<string> BlockLines(Accessibility accessibility)
        {
            var blockLines = new List<string>();

            blockLines.Add($"{accessibility.TokenString()} {Tokens.Type} {AsTypeName}");

            _members.ForEach(m => blockLines.Add($"{m.PropertyName} As {m.AsTypeName}"));

            blockLines.Add("End Type");

            return blockLines;
        }
    }
}
