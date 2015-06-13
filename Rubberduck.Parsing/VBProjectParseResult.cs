using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime.Tree;
using NetOffice.VBIDEApi;
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

                var listener = new IdentifierReferenceListener(componentParseResult, _declarations);
                var walker = new ParseTreeWalker();
                walker.Walk(listener, componentParseResult.ParseTree);
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