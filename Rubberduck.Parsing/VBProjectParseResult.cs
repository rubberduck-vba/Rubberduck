using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Parsing
{
    public class VBProjectParseResult
    {
        public VBProjectParseResult(VBProject project, IEnumerable<VBComponentParseResult> parseResults, IHostApplication hostApplication)
        {
            _project = project;
            _parseResults = parseResults;
            _declarations = new Declarations();

            var projectIdentifier = project.Name;
            var memberName = new QualifiedMemberName(new QualifiedModuleName(project), projectIdentifier);
            var projectDeclaration = new Declaration(memberName, null, null, projectIdentifier, false, false, Accessibility.Global, DeclarationType.Project, false);
            _declarations.Add(projectDeclaration);

            foreach (var declaration in VbaStandardLib.Declarations)
            {
                declaration.ClearReferences();
                _declarations.Add(declaration);
            }

            if (hostApplication != null && hostApplication.ApplicationName == "Excel")
            {
                foreach (var declaration in ExcelObjectModel.Declarations)
                {
                    declaration.ClearReferences();
                    _declarations.Add(declaration);
                }
            }

            if (project.References != null && project.References.Cast<Reference>().Any(r => r.Name == "ADODB"))
            {
                foreach (var declaration in AdodbObjectModel.Declarations)
                {
                    declaration.ClearReferences();
                    _declarations.Add(declaration);
                }
            }

            foreach (var declaration in _parseResults.SelectMany(item => item.Declarations))
            {
                _declarations.Add(declaration);
            }
        }

        public event EventHandler<ResolutionProgressEventArgs> Progress;

        private void OnProgress(ResolutionProgressEventArgs args)
        {
            var handler = Progress;
            if (handler != null)
            {
                handler(this, args);
            }
        }

        public void Resolve()
        {
            foreach (var componentParseResult in _parseResults)
            {
                var component = componentParseResult;
                Resolve(component);
            }
        }

        public async Task ResolveAsync()
        {
            foreach (var componentParseResult in _parseResults)
            {
                var component = componentParseResult;
                Task.Run(() => Resolve(component));
            }
        }

        private void Resolve(VBComponentParseResult componentParseResult)
        {
            try
            {
                var memberCount = componentParseResult.Declarations.Count(item => item.DeclarationType.HasFlag(DeclarationType.Member));
                var processedCount = 0;

                var resolver = new IdentifierReferenceResolver(componentParseResult.QualifiedName, _declarations);
                var listener = new IdentifierReferenceListener(resolver);
                listener.MemberProcessed += delegate
                {
                    processedCount++;
                    OnProgress(new ResolutionProgressEventArgs(componentParseResult.Component, (decimal)processedCount / memberCount));
                };

                var walker = new ParseTreeWalker();
                walker.Walk(listener, componentParseResult.ParseTree);
            }
            catch (InvalidOperationException)
            {
                // could not resolve all identifier references in this module.
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