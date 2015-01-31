using System.Collections.Generic;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections;

namespace Rubberduck.VBA.Nodes
{
    public class VBProjectParseResult
    {
        private readonly IDictionary<QualifiedModuleName, IParseTree> _results;

        public VBProjectParseResult()
            : this(new Dictionary<QualifiedModuleName, IParseTree>())
        { }

        public VBProjectParseResult(IDictionary<QualifiedModuleName,IParseTree> results)
        {
            _results = results;
        }

        public IParseTree this[QualifiedModuleName qualifiedName]
        {
            get { return _results[qualifiedName]; }
        }

        public IEnumerable<QualifiedModuleName> QualifiedModuleNames { get { return _results.Keys; } } 

        public void AddParseTree(QualifiedModuleName qualifiedName, IParseTree parseTree)
        {
            _results.Add(qualifiedName, parseTree);
        }
    }
}