using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;

namespace Rubberduck.Parsing.VBA
{
    public class ModuleState
    {
        public ConcurrentDictionary<Declaration, byte> Declarations { get; private set; }
        public ConcurrentDictionary<UnboundMemberDeclaration, byte> UnresolvedMemberDeclarations { get; private set; }
        public ITokenStream CodePaneTokenStream { get; private set; }
        public ITokenStream AttributesTokenStream { get; private set; }
        public IParseTree ParseTree { get; private set; }
        public IParseTree AttributesPassParseTree { get; private set; }
        public ParserState State { get; private set; }
        public int ModuleContentHashCode { get; private set; }
        public List<CommentNode> Comments { get; private set; }
        public List<IAnnotation> Annotations { get; private set; }
        public SyntaxErrorException ModuleException { get; private set; }
        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> ModuleAttributes { get; private set; }
        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> MembersAllowingAttributes { get; private set; }

        public bool IsNew { get; private set; }
        public bool IsMarkedAsModified { get; private set; }


        public ModuleState(ConcurrentDictionary<Declaration, byte> declarations)
        {
            Declarations = declarations;
            UnresolvedMemberDeclarations = new ConcurrentDictionary<UnboundMemberDeclaration, byte>();
            ParseTree = null;

            ModuleContentHashCode = 0;
            Comments = new List<CommentNode>();
            Annotations = new List<IAnnotation>();
            ModuleException = null;
            ModuleAttributes = new Dictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes>();
            MembersAllowingAttributes = new Dictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext>();

            IsNew = true;
            IsMarkedAsModified = false;
            State = ParserState.Pending;
        }

        public ModuleState(ParserState state)
        {
            Declarations = new ConcurrentDictionary<Declaration, byte>();
            UnresolvedMemberDeclarations = new ConcurrentDictionary<UnboundMemberDeclaration, byte>();
            ParseTree = null;
            State = state;
            ModuleContentHashCode = 0;
            Comments = new List<CommentNode>();
            Annotations = new List<IAnnotation>();
            ModuleException = null;
            ModuleAttributes = new Dictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes>();
            MembersAllowingAttributes = new Dictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext>();

            IsNew = true;
        }

        public ModuleState(SyntaxErrorException moduleException)
        {
            Declarations = new ConcurrentDictionary<Declaration, byte>();
            UnresolvedMemberDeclarations = new ConcurrentDictionary<UnboundMemberDeclaration, byte>();
            ParseTree = null;
            State = ParserState.Error;
            ModuleContentHashCode = 0;
            Comments = new List<CommentNode>();
            Annotations = new List<IAnnotation>();
            ModuleException = moduleException;
            ModuleAttributes = new Dictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes>();
            MembersAllowingAttributes = new Dictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext>();

            IsNew = true;
        }

        public ModuleState SetCodePaneTokenStream(ITokenStream codePaneTokenStream)
        {
            CodePaneTokenStream = codePaneTokenStream;
            return this;
        }

        public ModuleState SetParseTree(IParseTree parseTree, CodeKind codeKind)
        {
            switch (codeKind)
            {
                case CodeKind.AttributesCode:
                    AttributesPassParseTree = parseTree;
                    break;
                case CodeKind.CodePaneCode:
                    ParseTree = parseTree;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(codeKind), codeKind, null);
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

        public ModuleState SetModuleAttributes(IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> moduleAttributes)
        {
            ModuleAttributes = moduleAttributes;
            return this;
        }

        public ModuleState SetMembersAllowingAttributes(IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> membersAllowingAttributes)
        {
            MembersAllowingAttributes = membersAllowingAttributes;
            return this;
        }

        public ModuleState SetAttributesTokenStream(ITokenStream attributesTokenStream)
        {
            AttributesTokenStream = attributesTokenStream;
            return this;
        }

        public void MarkAsModified()
        {
            IsMarkedAsModified = true;
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