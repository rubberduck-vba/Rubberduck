using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Parsing.VBA
{
    public enum ResolutionState
    {
        Unresolved,
        InProgress,
        Resolved
    }

    public interface IRubberduckParserFactory
    {
        IRubberduckParser Create();
    }

    public class RubberduckParser : IRubberduckParser
    {
        private static readonly ConcurrentDictionary<QualifiedModuleName, VBComponentParseResult> ParseResultCache = 
            new ConcurrentDictionary<QualifiedModuleName, VBComponentParseResult>();

        private IHostApplication _hostApplication;

        private readonly Logger _logger;

        public RubberduckParser()
        {
            _logger = LogManager.GetCurrentClassLogger();
        }

        private readonly RubberduckParserState _state = new RubberduckParserState();
        public RubberduckParserState State { get { return _state; } }

        private IEnumerable<CommentNode> ParseComments(QualifiedModuleName qualifiedName)
        {
            var code = qualifiedName.Component.CodeModule.Code();
            var commentBuilder = new StringBuilder();
            var continuing = false;

            var startLine = 0;
            var startColumn = 0;

            for (var i = 0; i < code.Length; i++)
            {
                var line = code[i];                
                var index = 0;

                if (continuing || line.HasComment(out index))
                {
                    startLine = continuing ? startLine : i;
                    startColumn = continuing ? startColumn : index;

                    var commentLength = line.Length - index;

                    continuing = line.EndsWith("_");
                    if (!continuing)
                    {
                        commentBuilder.Append(line.Substring(index, commentLength).TrimStart());
                        var selection = new Selection(startLine + 1, startColumn + 1, i + 1, line.Length + 1);

                        var result = new CommentNode(commentBuilder.ToString(), new QualifiedSelection(qualifiedName, selection));
                        commentBuilder.Clear();
                        
                        yield return result;
                    }
                    else
                    {
                        // ignore line continuations in comment text:
                        commentBuilder.Append(line.Substring(index, commentLength).TrimStart()); 
                    }
                }
            }
        }

        private static IParseTree Parse(string code, IParseTreeListener listener, out ITokenStream outStream)
        {
            var input = new AntlrInputStream(code);
            var lexer = new VBALexer(input);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);

            parser.AddErrorListener(new ExceptionErrorListener());
            parser.AddParseListener(listener);

            outStream = tokens;

            var result = parser.startRule();
            return result;
        }

        private async Task<Tuple<IParseTree, ITokenStream, Action>> ParseAsync(string code, IParseTreeListener listener)
        {
            return await Task.Run(() =>
            {
                ITokenStream stream;
                var tree = Parse(code, listener, out stream);
                return Tuple.Create(tree, stream, new Action(() => { ResolveReferences(tree); }));
            });
        }

        public void Parse(VBE vbe)
        {
            foreach (var task in vbe.VBProjects.Cast<VBProject>()
                .Select(project => new Task(() => Parse(project))))
            {
                task.Start();
            }
        }

        public void Parse(VBProject vbProject)
        {
            foreach (var task in vbProject.VBComponents.Cast<VBComponent>()
                .Select(component => new Task(() => ParseAsync(component))))
            {
                task.Start();
            }
        }

        public async Task ParseAsync(VBComponent vbComponent)
        {
            var qualifiedName = new QualifiedModuleName(vbComponent);
            var comments = ParseComments(qualifiedName);
            var listener = new DeclarationSymbolsListener(qualifiedName, Accessibility.Implicit, vbComponent.Type, comments);

            var code = vbComponent.CodeModule.Lines();
            var result = await ParseAsync(code, listener);

            var scope = vbComponent.Collection.Parent.Name + "." + vbComponent.Name;
            _state.MarkForResolution(scope);
            _state.AddTokenStream(vbComponent, result.Item2);            
        }

        private void ResolveReferences(IParseTree tree)
        {
            var tasks = _state.UnresolvedDeclarations
                .GroupBy(declaration => declaration.QualifiedSelection.QualifiedName)
                .Select(grouping => new Task(() =>
                {
                    var resolver = new IdentifierReferenceResolver(grouping.Key, grouping);
                    var listener = new IdentifierReferenceListener(resolver);
                    var walker = new ParseTreeWalker();
                    walker.Walk(listener, tree);
                }));
            foreach (var task in tasks)
            {
                task.Start();
            }
        }
    }
}
