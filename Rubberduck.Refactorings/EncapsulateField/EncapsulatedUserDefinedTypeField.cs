using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulatedUserDefinedTypeField : IEncapsulateFieldCandidate
    {
        IEnumerable<IEncapsulatedUserDefinedTypeMember> Members { get; }
        void AddMember(IEncapsulatedUserDefinedTypeMember member);
        bool FieldQualifyMemberPropertyNames { set; get; }
        bool TypeDeclarationIsPrivate { set; get; }
    }

    public class EncapsulatedUserDefinedTypeField : EncapsulateFieldCandidate, IEncapsulatedUserDefinedTypeField
    {
        public EncapsulatedUserDefinedTypeField(Declaration declaration, IEncapsulateFieldNamesValidator validator)
            : base(declaration, validator)
        {
            PropertyAccessExpression = () => EncapsulateFlag ? NewFieldName : IdentifierName;
        }

        public void AddMember(IEncapsulatedUserDefinedTypeMember member)
        {
            _udtMembers.Add(member);
        }

        private List<IEncapsulatedUserDefinedTypeMember> _udtMembers = new List<IEncapsulatedUserDefinedTypeMember>();
        public IEnumerable<IEncapsulatedUserDefinedTypeMember> Members => _udtMembers;

        public bool TypeDeclarationIsPrivate { set; get; }

        public override string NewFieldName
        {
            get => TypeDeclarationIsPrivate ? _fieldAndProperty.TargetFieldName : _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
        }

        public bool FieldQualifyMemberPropertyNames
        {
            set
            {
                foreach (var member in Members)
                {
                    member.FieldQualifyPropertyName = value;
                }
            }

            get => Members.All(m => m.FieldQualifyPropertyName);
        }

        public override void LoadReferenceExpressionChanges()
        {
            if (TypeDeclarationIsPrivate)
            {
                LoadPrivateUDTFieldReferenceExpressions();
                LoadUDTMemberReferenceExpressions();
                return;
            }
            LoadFieldReferenceExpressions();
        }

        public override IEnumerable<IPropertyGeneratorSpecification> PropertyGenerationSpecs
        {
            get
            {
                if (TypeDeclarationIsPrivate)
                {
                    var specs = new List<IPropertyGeneratorSpecification>();
                    foreach (var member in Members)
                    {
                        specs.Add(member.AsPropertyGeneratorSpec);
                    }
                    return specs;
                }
                return new List<IPropertyGeneratorSpecification>() { AsPropertyGeneratorSpec };
            }
        }

        public override IEnumerable<KeyValuePair<IdentifierReference, RewriteReplacePair>> ReferenceReplacements
        {
            get
            {
                var results = new List<KeyValuePair<IdentifierReference, RewriteReplacePair>>();
                foreach (var replacement in IdentifierReplacements)
                {
                    var kv = new KeyValuePair<IdentifierReference, RewriteReplacePair>
                        (replacement.Key, replacement.Value);
                    results.Add(kv);
                }
                foreach (var member in Members)
                {
                    foreach (var replacement in member.IdentifierReplacements)
                    {
                        var kv = new KeyValuePair<IdentifierReference, RewriteReplacePair>
                            (replacement.Key, replacement.Value);
                        results.Add(kv);
                    }
                }
                return results;
            }
        }

        private void LoadPrivateUDTFieldReferenceExpressions()
        {
            foreach (var idRef in References)
            {
                if (idRef.QualifiedModuleName == QualifiedModuleName
                    && idRef.Context.Parent.Parent is VBAParser.WithStmtContext wsc)
                {
                    SetReferenceRewriteContent(idRef, NewFieldName);
                }
            }
        }

        private void LoadUDTMemberReferenceExpressions()
        {
            foreach (var member in Members)
            {
                foreach (var rf in member.References)
                {
                    if (rf.QualifiedModuleName == QualifiedModuleName
                        && !rf.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out _))
                    {
                        member.SetReferenceRewriteContent(rf, member.PropertyName);
                    }
                    else
                    {
                        //If rf is a WithMemberAccess expression, modify the LExpr.  e.g. "With this" => "With <qmn.ModuleName>"
                        var moduleQualifier = rf.Context.TryGetAncestor<VBAParser.WithStmtContext>(out _)
                            || rf.QualifiedModuleName == QualifiedModuleName
                            ? string.Empty
                            : $"{QualifiedModuleName.ComponentName}";

                        member.SetReferenceRewriteContent(rf, $"{moduleQualifier}.{member.PropertyName}");
                    }
                }
            }
        }
    }
}
