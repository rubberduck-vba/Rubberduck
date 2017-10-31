using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public class ModuleState
    {
        public ConcurrentDictionary<Declaration, byte> Declarations { get; private set; }
        public ConcurrentDictionary<UnboundMemberDeclaration, byte> UnresolvedMemberDeclarations { get; private set; }
        public ITokenStream TokenStream { get; private set; }
        public IModuleRewriter ModuleRewriter { get; private set; }
        public IModuleRewriter AttributesRewriter { get; private set; }
        public IParseTree ParseTree { get; private set; }
        public IParseTree AttributesPassParseTree { get; private set; }
        public ParserState State { get; private set; }
        public int ModuleContentHashCode { get; private set; }
        public List<CommentNode> Comments { get; private set; }
        public List<IAnnotation> Annotations { get; private set; }
        public SyntaxErrorException ModuleException { get; private set; }
        public IDictionary<Tuple<string, DeclarationType>, Attributes> ModuleAttributes { get; private set; }

        public bool IsNew { get; private set; }

        public ModuleState(ConcurrentDictionary<Declaration, byte> declarations)
        {
            Declarations = declarations;
            UnresolvedMemberDeclarations = new ConcurrentDictionary<UnboundMemberDeclaration, byte>();
            TokenStream = null;
            ParseTree = null;

            ModuleContentHashCode = 0;
            Comments = new List<CommentNode>();
            Annotations = new List<IAnnotation>();
            ModuleException = null;
            ModuleAttributes = new Dictionary<Tuple<string, DeclarationType>, Attributes>();

            IsNew = true;
            State = ParserState.Pending;
        }

        public ModuleState(ParserState state)
        {
            Declarations = new ConcurrentDictionary<Declaration, byte>();
            UnresolvedMemberDeclarations = new ConcurrentDictionary<UnboundMemberDeclaration, byte>();
            TokenStream = null;
            ParseTree = null;
            State = state;
            ModuleContentHashCode = 0;
            Comments = new List<CommentNode>();
            Annotations = new List<IAnnotation>();
            ModuleException = null;
            ModuleAttributes = new Dictionary<Tuple<string, DeclarationType>, Attributes>();

            IsNew = true;
        }

        public ModuleState(SyntaxErrorException moduleException)
        {
            Declarations = new ConcurrentDictionary<Declaration, byte>();
            UnresolvedMemberDeclarations = new ConcurrentDictionary<UnboundMemberDeclaration, byte>();
            TokenStream = null;
            ParseTree = null;
            State = ParserState.Error;
            ModuleContentHashCode = 0;
            Comments = new List<CommentNode>();
            Annotations = new List<IAnnotation>();
            ModuleException = moduleException;
            ModuleAttributes = new Dictionary<Tuple<string, DeclarationType>, Attributes>();

            IsNew = true;
        }

        public ModuleState(IDictionary<Tuple<string, DeclarationType>, Attributes> moduleAttributes)
        {
            Declarations = new ConcurrentDictionary<Declaration, byte>();
            UnresolvedMemberDeclarations = new ConcurrentDictionary<UnboundMemberDeclaration, byte>();
            TokenStream = null;
            ParseTree = null;
            State = ParserState.None;
            ModuleContentHashCode = 0;
            Comments = new List<CommentNode>();
            Annotations = new List<IAnnotation>();
            ModuleException = null;
            ModuleAttributes = moduleAttributes;

            IsNew = true;
        }

        public ModuleState SetTokenStream(ICodeModule module, ITokenStream tokenStream)
        {
            TokenStream = tokenStream;
            var tokenStreamRewriter = new TokenStreamRewriter(tokenStream);
            ModuleRewriter = new ModuleRewriter(module, tokenStreamRewriter);
            return this;
        }

        public ModuleState SetParseTree(IParseTree parseTree, ParsePass pass)
        {
            switch (pass)
            {
                case ParsePass.AttributesPass:
                    AttributesPassParseTree = parseTree;
                    break;
                case ParsePass.CodePanePass:
                    ParseTree = parseTree;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(pass), pass, null);
            }
            return this;
        }

        public ModuleState SetState(ParserState state)
        {
            State = state;
            return this;
        }

        public ModuleState SetModuleContentHashCode(int moduleContentHashCode)
        {
            ModuleContentHashCode = moduleContentHashCode;
            IsNew = false;
            return this;
        }

        public ModuleState SetComments(List<CommentNode> comments)
        {
            Comments = comments;
            return this;
        }

        public ModuleState SetAnnotations(List<IAnnotation> annotations)
        {
            Annotations = annotations;
            return this;
        }

        public ModuleState SetModuleException(SyntaxErrorException moduleException)
        {
            ModuleException = moduleException;
            return this;
        }

        public ModuleState SetModuleAttributes(IDictionary<Tuple<string, DeclarationType>, Attributes> moduleAttributes)
        {
            ModuleAttributes = moduleAttributes;
            return this;
        }

        public ModuleState SetAttributesRewriter(IModuleRewriter rewriter)
        {
            AttributesRewriter = rewriter;
            return this;
        }

        private bool _isDisposed;

        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            Declarations?.Clear();
            Comments?.Clear();
            Annotations?.Clear();
            ModuleAttributes?.Clear();

            _isDisposed = true;
        }
    }
}