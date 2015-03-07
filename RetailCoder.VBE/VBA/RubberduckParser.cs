using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.VBA
{
    public class RubberduckParser : IRubberduckParser
    {
        private static readonly ConcurrentDictionary<QualifiedModuleName, VBComponentParseResult> ParseResultCache = 
            new ConcurrentDictionary<QualifiedModuleName, VBComponentParseResult>();

        public IParseTree Parse(string code)
        {
            var input = new AntlrInputStream(code);
            var lexer = new VBLexer(input);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBParser(tokens);
            
            var result = parser.StartRule();
            return result;
        }

        public IEnumerable<VBComponentParseResult> Parse(VBProject project)
        {
            if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
            {
                yield break;
            }

            var modules = project.VBComponents.Cast<VBComponent>();
            foreach(var module in modules)
            {
                yield return Parse(module);
            };
        }

        public VBComponentParseResult Parse(VBComponent component)
        {
            VBComponentParseResult cachedValue;
            var name = component.QualifiedName();
            if (ParseResultCache.TryGetValue(name, out cachedValue))
            {
                return cachedValue;
            }

            var parseTree = Parse(component.CodeModule.Lines());
            var comments = ParseComments(component);
            var result = new VBComponentParseResult(component, parseTree, comments);

            ParseResultCache.AddOrUpdate(name, module => result, (qName, module) => result);
            return result;
        }

        public IEnumerable<CommentNode> ParseComments(VBComponent component)
        {
            var code = component.CodeModule.Code();
            var qualifiedName = component.QualifiedName();

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
