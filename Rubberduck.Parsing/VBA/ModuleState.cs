using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public class ModuleState
    {
        public ConcurrentDictionary<Declaration, byte> Declarations { get; private set; }
        public ConcurrentDictionary<UnboundMemberDeclaration, byte> UnresolvedMemberDeclarations { get; private set; }
        public ITokenStream TokenStream { get; private set; }
        public TokenStreamRewriter Rewriter { get; private set; }
        public IParseTree ParseTree { get; private set; }
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

            if (declarations.Any() && declarations.ElementAt(0).Key.QualifiedName.QualifiedModuleName.Component != null)
            {
                State = ParserState.Pending;
            }
            else
            {
                State = ParserState.Pending;
            }

            ModuleContentHashCode = 0;
            Comments = new List<CommentNode>();
            Annotations = new List<IAnnotation>();
            ModuleException = null;
            ModuleAttributes = new Dictionary<Tuple<string, DeclarationType>, Attributes>();

            IsNew = true;
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

        public ModuleState SetTokenStream(ITokenStream tokenStream)
        {
            TokenStream = tokenStream;
            Rewriter = new TokenStreamRewriter(tokenStream);
            return this;
        }

        public ModuleState SetParseTree(IParseTree parseTree)
        {
            ParseTree = parseTree;
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


        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            if (Declarations != null)
            {
                Declarations.Clear();
            }

            if (Comments != null)
            {
                Comments.Clear();
            }

            if (Annotations != null)
            {
                Annotations.Clear();
            }

            if (ModuleAttributes != null)
            {
                ModuleAttributes.Clear();
            }

            _isDisposed = true;
        }
    }
}