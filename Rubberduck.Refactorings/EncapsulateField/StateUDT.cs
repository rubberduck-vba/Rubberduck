using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Common;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IStateUDT
    {
        string TypeIdentifier { set; get; }
        string FieldIdentifier { set; get; }
        string TypeDeclarationBlock(IIndenter indenter = null);
        string FieldDeclarationBlock { get; }
        void AddMembers(IEnumerable<IEncapsulateFieldCandidate> fields);
        QualifiedModuleName QualifiedModuleName { set; get; }
    }

    /*
     * StateUDT is the UserDefinedType introduced by this refactoring 
     * whose members represent object state in lieu of individually declared member variables/fields.
     */
    public class StateUDT : IStateUDT
    {
        private static string _defaultNewFieldName = EncapsulateFieldResources.DefaultStateUDTFieldName;
        private List<IEncapsulateFieldCandidate> _members;
        private readonly IEncapsulateFieldNamesValidator _validator;

        public StateUDT(QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
            :this($"{EncapsulateFieldResources.StateUserDefinedTypeIdentifierPrefix}{qmn.ComponentName.CapitalizeFirstLetter()}", validator)
        {
            QualifiedModuleName = qmn;
        }

        public StateUDT(string typeIdentifier, IEncapsulateFieldNamesValidator validator)
        {
            _validator = validator;
            FieldIdentifier = _defaultNewFieldName;
            TypeIdentifier = typeIdentifier;
            _members = new List<IEncapsulateFieldCandidate>();
        }

        public QualifiedModuleName QualifiedModuleName { set;  get; }

        public string TypeIdentifier { set; get; }

        public string FieldIdentifier { set; get; }

        public void AddMembers(IEnumerable<IEncapsulateFieldCandidate> fields) => _members = fields.ToList();

        public string FieldDeclarationBlock
            => $"{Accessibility.Private} {FieldIdentifier} {Tokens.As} {TypeIdentifier}";

        public string TypeDeclarationBlock(IIndenter indenter = null)
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

            blockLines.Add($"{accessibility.TokenString()} {Tokens.Type} {TypeIdentifier}");

            _members.ForEach(m => blockLines.Add($"{m.AsUDTMemberDeclaration}"));

            blockLines.Add($"{Tokens.End} {Tokens.Type}");

            return blockLines;
        }
    }
}
