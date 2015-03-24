using System.Collections.Generic;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing
{
    public class VBProjectParseResult
    {
        public VBProjectParseResult(IEnumerable<VBComponentParseResult> parseResults)
        {
            _parseResults = parseResults;
            _declarations = new Declarations();
            IdentifySymbols();
            IdentifySymbolUsages();
        }

        /// <summary>
        /// Locates all declared symbols (identifiers) in the project.
        /// </summary>
        /// <remarks>
        /// This method walks the entire parse tree for each module.
        /// </remarks>
        private void IdentifySymbols()
        {
            foreach (var componentParseResult in _parseResults)
            {
                var listener = new DeclarationSymbolsListener(componentParseResult);
                var walker = new ParseTreeWalker();
                walker.Walk(listener, componentParseResult.ParseTree);

                foreach (var declaration in listener.Declarations.Items)
                {
                    _declarations.Add(declaration);
                }
            }
        }

        private void IdentifySymbolUsages()
        {
            foreach (var componentParseResult in _parseResults)
            {
                var listener = new IdentifierReferenceListener(componentParseResult, _declarations);
                var walker = new ParseTreeWalker();
                walker.Walk(listener, componentParseResult.ParseTree);
            }
        }

        private readonly IEnumerable<VBComponentParseResult> _parseResults;
        
        private readonly Declarations _declarations;
        public Declarations Declarations { get { return _declarations; } }

        public IEnumerable<VBComponentParseResult> ComponentParseResults { get { return _parseResults; } }
    }
}