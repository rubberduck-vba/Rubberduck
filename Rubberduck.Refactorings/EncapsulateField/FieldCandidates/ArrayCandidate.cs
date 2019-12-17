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
    public class ArrayCandidate : EncapsulateFieldCandidate
    {
        private string _subscripts;
        public ArrayCandidate(Declaration declaration, IEncapsulateFieldNamesValidator validator)
            :base(declaration, validator)
        {
            ImplementLetSetterType = false;
            ImplementSetSetterType = false;
            AsTypeName = declaration.AsTypeName;
            AsTypeName_Property = Tokens.Variant;
            CanBeReadWrite = false;
            IsReadOnly = true;

            _subscripts = string.Empty;
            if (declaration.Context.TryGetChildContext<VBAParser.SubscriptsContext>(out var ctxt))
            {
                _subscripts = ctxt.GetText();
            }
        }

        public override string AsUDTMemberDeclaration
            => $"{PropertyName}({_subscripts}) {Tokens.As} {AsTypeName}";
    }
}
