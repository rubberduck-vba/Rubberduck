using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    public class VBProjectParseResult
    {
        public VBProjectParseResult(VBProject project, IEnumerable<VBComponentParseResult> parseResults)
        {
            _project = project;
            _parseResults = parseResults;
            _declarations = new Declarations();

            var projectIdentifier = project.Name;
            var memberName = new QualifiedMemberName(new QualifiedModuleName(project), projectIdentifier);
            var projectDeclaration = new Declaration(memberName, "VBE", projectIdentifier, false, false, Accessibility.Global, DeclarationType.Project, false);
            _declarations.Add(projectDeclaration);

            foreach (var declaration in VbaStandardLib.Declarations)
            {
                _declarations.Add(declaration);
            }

            foreach (var declaration in _parseResults.SelectMany(item => item.Declarations))
            {
                _declarations.Add(declaration);
            }
        }

        public event EventHandler<ResolutionProgressEventArgs> Progress;

        private void OnProgress(VBComponentParseResult result)
        {
            var handler = Progress;
            if (handler != null)
            {
                handler(null, new ResolutionProgressEventArgs(result.Component));
            }
        }

        public void Resolve()
        {
            foreach (var componentParseResult in _parseResults)
            {
                OnProgress(componentParseResult);

                try
                {
                    var resolver = new IdentifierReferenceResolver(componentParseResult.QualifiedName, _declarations);
                    var listener = new IdentifierReferenceListener(resolver);
                    var walker = new ParseTreeWalker();
                    walker.Walk(listener, componentParseResult.ParseTree);
                }
                catch (InvalidOperationException)
                {
                    // could not resolve all identifier references in this module.
                }
            }
        }

        private readonly IEnumerable<VBComponentParseResult> _parseResults;
        
        private readonly Declarations _declarations;
        public Declarations Declarations { get { return _declarations; } }

        public IEnumerable<VBComponentParseResult> ComponentParseResults { get { return _parseResults; } }

        private readonly VBProject _project;
        public VBProject Project { get { return _project; } }
    }
}