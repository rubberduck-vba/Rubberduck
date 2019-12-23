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
    public interface IObjectStateUDT : IEncapsulateFieldDeclaration
    {
        string TypeIdentifier { set; get; }
        string FieldIdentifier { set; get; }
        string TypeDeclarationBlock(IIndenter indenter = null);
        string FieldDeclarationBlock { get; }
        void AddMembers(IEnumerable<IEncapsulateFieldCandidate> fields);
        bool IsExistingDeclaration { get; }
    }

    /*
     * StateUDT is the UserDefinedType introduced by this refactoring 
     * whose members represent object state in lieu of individually declared member variables/fields.
     */
    public class ObjectStateUDT : IObjectStateUDT//, IEncapsulateFieldDeclaration
    {
        private static string _defaultNewFieldName = EncapsulateFieldResources.DefaultStateUDTFieldName;
        private List<IEncapsulateFieldCandidate> _members;
        private readonly IUserDefinedTypeCandidate _decoratedUDT;

        public ObjectStateUDT(IUserDefinedTypeCandidate udt)
            :this(udt.Declaration.AsTypeName)
        {
            if (!udt.Declaration.Accessibility.Equals(Accessibility.Private))
            {
                throw new ArgumentException();
            }

            _decoratedUDT = udt;
            QualifiedModuleName = udt.Declaration.QualifiedModuleName;
        }

        public ObjectStateUDT(QualifiedModuleName qmn)
            :this($"{EncapsulateFieldResources.StateUserDefinedTypeIdentifierPrefix}{qmn.ComponentName.CapitalizeFirstLetter()}")
        {
            QualifiedModuleName = qmn;
        }

        public ObjectStateUDT(string typeIdentifier)
        {
            FieldIdentifier = _defaultNewFieldName;
            TypeIdentifier = typeIdentifier;
            _members = new List<IEncapsulateFieldCandidate>();
        }

        public string IdentifierName => _decoratedUDT?.IdentifierName ?? FieldIdentifier;

        public Selection Selection => _decoratedUDT?.Selection ?? new Selection();

        public Accessibility Accessibility => Accessibility.Private; // _decoratedUDT.Accessibility;

        public string AsTypeName => _decoratedUDT?.AsTypeName ?? TypeIdentifier;

        public QualifiedModuleName QualifiedModuleName { set;  get; }

        public string TypeIdentifier { set; get; }

        public bool IsExistingDeclaration => _decoratedUDT != null;

        public string FieldIdentifier { set; get; }

        public void AddMembers(IEnumerable<IEncapsulateFieldCandidate> fields)
        {
            if (IsExistingDeclaration)
            {
                _members = _decoratedUDT.Members.Select(m => m).Cast<IEncapsulateFieldCandidate>().ToList();
            }
            _members.AddRange(fields);
        }

        public string FieldDeclarationBlock
            => $"{Accessibility.Private} {IdentifierName} {Tokens.As} {AsTypeName}";

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
