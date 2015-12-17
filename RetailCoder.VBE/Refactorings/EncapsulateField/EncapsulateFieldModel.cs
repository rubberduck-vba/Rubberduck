using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel
    {
        private readonly RubberduckParserState _parseResult;
        public RubberduckParserState ParseResult { get { return _parseResult; } }

        private readonly IList<Declaration> _declarations;
        public IEnumerable<Declaration> Declarations { get { return _declarations; } }

        public Declaration TargetDeclaration { get; private set; }

        private readonly IMessageBox _messageBox;

        public EncapsulateFieldModel(RubberduckParserState parseResult, QualifiedSelection selection, IMessageBox messageBox)
        {
            _parseResult = parseResult;
            _declarations = parseResult.AllDeclarations.ToList();
            _messageBox = messageBox;
        }
    }
}