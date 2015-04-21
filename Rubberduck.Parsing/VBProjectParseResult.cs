using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
            foreach (var componentParseResult in _parseResults.Where(r => r.Component != null))
            {
                try
                {
                    var listener = new DeclarationSymbolsListener(componentParseResult);
                    var walker = new ParseTreeWalker();
                    walker.Walk(listener, componentParseResult.ParseTree);

                    if (!_declarations.Items.Any())
                    { 
                        var projectIdentifier = componentParseResult.QualifiedName.Project.Name;
                        var memberName = componentParseResult.QualifiedName.QualifyMemberName(projectIdentifier);
                        var projectDeclaration = new Declaration(memberName, "VBE", projectIdentifier, projectIdentifier, false, false, Accessibility.Global, DeclarationType.Project, null, Selection.Home);
                        _declarations.Add(projectDeclaration);
                    }

                    foreach (var declaration in listener.Declarations.Items)
                    {
                        _declarations.Add(declaration);
                    }
                }
                catch (COMException)
                {
                    // something happened, couldn't access VBComponent for some reason
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