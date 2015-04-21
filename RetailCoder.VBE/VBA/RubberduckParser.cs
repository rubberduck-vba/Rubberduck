using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;

namespace Rubberduck.VBA
{
    internal class RubberduckParser : IRubberduckParser
    {
        private static readonly ConcurrentDictionary<QualifiedModuleName, VBComponentParseResult> ParseResultCache = 
            new ConcurrentDictionary<QualifiedModuleName, VBComponentParseResult>();

        public VBProjectParseResult Parse(VBProject project)
        {
            var results = new List<VBComponentParseResult>();
            if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
            {
                return new VBProjectParseResult(results);
            }

            var modules = project.VBComponents.Cast<VBComponent>();
            results.AddRange(modules.Select(Parse).Where(result => result != null));

            return new VBProjectParseResult(results);
        }

        private IParseTree Parse(string code, out TokenStreamRewriter outRewriter)
        {
            var input = new AntlrInputStream(code);
            var lexer = new VBALexer(input);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            outRewriter = new TokenStreamRewriter(tokens);

            var result = parser.startRule();
            return result;
        }

        private VBComponentParseResult Parse(VBComponent component)
        {
            try
            {
                VBComponentParseResult cachedValue;
                var name = new QualifiedModuleName(component); // already a performance hit
                if (ParseResultCache.TryGetValue(name, out cachedValue))
                {
                    return cachedValue;
                }

                var codeModule = component.CodeModule;
                var lines = codeModule.Lines();

                TokenStreamRewriter rewriter;
                var parseTree = Parse(lines, out rewriter);
                var comments = ParseComments(name);
                var result = new VBComponentParseResult(component, parseTree, comments, rewriter);

                ParseResultCache.AddOrUpdate(name, module => result, (qName, module) => result);
                return result;
            }
            catch (COMException)
            {
                return null;
            }
        }

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
                        var selection = new Selection(startLine + 1, startColumn + 1, i + 1, line.Length);

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
    }
}
