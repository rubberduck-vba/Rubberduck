using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Logging;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Parsing.VBA
{
    public class RubberduckParser : IRubberduckParser
    {
        private static readonly ConcurrentDictionary<QualifiedModuleName, VBComponentParseResult> ParseResultCache = 
            new ConcurrentDictionary<QualifiedModuleName, VBComponentParseResult>();

        private static IRubberduckFactory<IRubberduckCodePane> _factory;

        private static bool _isParsing;

        private readonly Logger _logger;

        public RubberduckParser(IRubberduckFactory<IRubberduckCodePane> factory)
        {
#if DEBUG
            LoggingConfigurator.ConfigureParserLogger();
#endif
            _logger = LogManager.GetCurrentClassLogger();

            _factory = factory;
            
        }

        public void RemoveProject(VBProject project)
        {
            foreach (var key in ParseResultCache.Keys.Where(k => k.Project.Equals(project)))
            {
                VBComponentParseResult result;
                ParseResultCache.TryRemove(key, out result);
            }
        }

        public VBProjectParseResult Parse(VBProject project, object owner = null)
        {
            if (owner != null)
            {
                OnParseStarted(new[]{project.Name}, owner);
            }

            var results = new List<VBComponentParseResult>();
            if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
            {
                return new VBProjectParseResult(project, results, _factory);
            }

            var modules = project.VBComponents.Cast<VBComponent>();
            var mustResolve = false;
            foreach (var vbComponent in modules)
            {
                OnParseProgress(vbComponent);

                bool fromCache;
                var componentResult = Parse(vbComponent, out fromCache);

                if (componentResult != null)
                {
                    mustResolve = mustResolve || !fromCache;
                    results.Add(componentResult);
                }
            }

            var parseResult = new VBProjectParseResult(project, results, _factory);
            if (mustResolve)
            {
                parseResult.Progress += parseResult_Progress;
                parseResult.Resolve();
                parseResult.Progress -= parseResult_Progress;
            }
            if (owner != null)
            {
                OnParseCompleted(new[] {parseResult}, owner);
            }

            return parseResult;
        }

        private void parseResult_Progress(object sender, ResolutionProgressEventArgs e)
        {
            OnResolveProgress(e.Component);
        }

        public IParseTree Parse(string code, out ITokenStream outStream)
        {
            var input = new AntlrInputStream(code);
            var lexer = new VBALexer(input);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            outStream = tokens;

            var result = parser.startRule();
            return result;
        }

        private VBComponentParseResult Parse(VBComponent component, out bool cached)
        {
            try
            {
                VBComponentParseResult cachedValue;
                var name = new QualifiedModuleName(component); // already a performance hit
                if (ParseResultCache.TryGetValue(name, out cachedValue))
                {
                    cached = true;
                    return cachedValue;
                }

                var codeModule = component.CodeModule;
                var lines = codeModule.Lines();

                ITokenStream stream;
                var parseTree = Parse(lines, out stream);
                var comments = ParseComments(name);
                var result = new VBComponentParseResult(component, parseTree, comments, stream, _factory);

                var existing = ParseResultCache.Keys.SingleOrDefault(k => k.Project == name.Project && k.ComponentName == name.ComponentName);
                VBComponentParseResult removed;
                ParseResultCache.TryRemove(existing, out removed);
                ParseResultCache.AddOrUpdate(name, module => result, (qName, module) => result);

                cached = false;
                return result;
            }
            catch (SyntaxErrorException exception)
            {
                OnParserError(exception, component);
                cached = false;
                return null;
            }
            catch (COMException)
            {
                cached = false;
                return null;
            }
        }

        public event EventHandler<ParseErrorEventArgs> ParserError;

        private void OnParserError(SyntaxErrorException exception, VBComponent component)
        {
            if (LogManager.IsLoggingEnabled())
            {
                LogParseException(exception, component);
            }

            var handler = ParserError;
            if (handler != null)
            {
                handler(this, new ParseErrorEventArgs(exception, component, new RubberduckCodePaneFactory()));
            }
        }

        private void LogParseException(SyntaxErrorException exception, VBComponent component)
        {
            var offendingProject = component.Collection.Parent.Name;
            var offendingComponent = component.Name;
            var offendingLine = component.CodeModule.get_Lines(exception.LineNumber, 1);

            var message = string.Format("Parser encountered a syntax error in {0}.{1}, line {2}. Content: '{3}'", offendingProject, offendingComponent, exception.LineNumber, offendingLine);
            _logger.ErrorException(message, exception);
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
                        var selection = new Selection(startLine + 1, startColumn + 1, i + 1, line.Length + 1);

                        var result = new CommentNode(commentBuilder.ToString(), new QualifiedSelection(qualifiedName, selection, _factory));
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

        public event EventHandler<ParseStartedEventArgs> ParseStarted;
        private void OnParseStarted(IEnumerable<string> projectNames, object owner)
        {
            var handler = ParseStarted;
            if (handler != null)
            {
                handler(owner, new ParseStartedEventArgs(projectNames));
            }
        }

        public event EventHandler<ResolutionProgressEventArgs> ResolutionProgress;
        private void OnResolveProgress(VBComponent component)
        {
            var handler = ResolutionProgress;
            if (handler != null)
            {
                handler(this, new ResolutionProgressEventArgs(component));
            }
        }

        public event EventHandler<ParseProgressEventArgs> ParseProgress;
        private void OnParseProgress(VBComponent component)
        {
            var handler = ParseProgress;
            if (handler != null)
            {
                handler(this, new ParseProgressEventArgs(component));
            }
        }

        public event EventHandler<ParseCompletedEventArgs> ParseCompleted;
        private void OnParseCompleted(IEnumerable<VBProjectParseResult> results, object owner)
        {
            var handler = ParseCompleted;
            if (handler != null)
            {
                handler(owner, new ParseCompletedEventArgs(results));
            }

            _isParsing = false;
        }

        public void Parse(VBE vbe, object owner)
        {
            if (!_isParsing)
            {
                _isParsing = true;

                var projects = vbe.VBProjects.Cast<VBProject>().ToList();
                OnParseStarted(projects.Select(project => project.Name), owner);

                var results = projects.AsParallel().Select(project => Parse(project)).ToList();
                OnParseCompleted(results, owner);
            }
        }
    }
}
