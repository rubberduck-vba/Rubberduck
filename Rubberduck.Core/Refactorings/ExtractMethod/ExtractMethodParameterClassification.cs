using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodParameterClassification : IExtractMethodParameterClassification
    {
        // https://github.com/rubberduck-vba/Rubberduck/wiki/Extract-Method-Refactoring-%3A-Workings---Determining-what-params-to-move
        private readonly IEnumerable<IExtractMethodRule> _emRules;
        private List<Declaration> _byref;
        private List<Declaration> _byval;
        private List<Declaration> _declarationsToMove;
        private List<Declaration> _extractDeclarations;

        public ExtractMethodParameterClassification(IEnumerable<IExtractMethodRule> emRules)
        {
            _emRules = emRules;
            _byref = new List<Declaration>();
            _byval = new List<Declaration>();
            _declarationsToMove = new List<Declaration>();
            _extractDeclarations = new List<Declaration>();
        }

        public void classifyDeclarations(QualifiedSelection selection, Declaration item)
        {

            byte flags = new Byte();
            foreach (var oRef in item.References)
            {
                foreach (var rule in _emRules)
                {
                    var byteFlag = rule.setValidFlag(oRef, selection.Selection);
                    flags = (byte)(flags | (byte)byteFlag);

                }
            }

            if (flags < 4) { /*ignore the variable*/ }
            else if (flags < 12)
                _byref.Add(item);
            else if (flags == 12)
                _declarationsToMove.Add(item);
            else if (flags > 12)
                _byval.Add(item);

            if (flags >= 18)
            {
                _extractDeclarations.Add(item);
            }
        }

        public IEnumerable<ExtractedParameter> ExtractedParameters
        {
            get {
                return _byref.Select(dec => new ExtractedParameter(dec.AsTypeName, ExtractedParameter.PassedBy.ByRef, dec.IdentifierName)).
                    Union(_byval.Select(dec => new ExtractedParameter(dec.AsTypeName, ExtractedParameter.PassedBy.ByVal, dec.IdentifierName)));
            }
        }

        public IEnumerable<Declaration> DeclarationsToMove { get { return _declarationsToMove; } }
        public IEnumerable<Declaration> ExtractedDeclarations { get { return _extractDeclarations; } }

    }
}
