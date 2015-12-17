using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Parsing.VBA
{
    public interface IRubberduckParserState {
        event EventHandler StateChanged;
        ParserState Status { get; }

        /// <summary>
        /// Gets <see cref="ParserRuleContext"/> objects representing 'Call' statements in the parse tree.
        /// </summary>
        IEnumerable<QualifiedContext> ObsoleteCallContexts { get; set; }

        /// <summary>
        /// Gets <see cref="ParserRuleContext"/> objects representing explicit 'Let' statements in the parse tree.
        /// </summary>
        IEnumerable<QualifiedContext> ObsoleteLetContexts { get; set; }

        IEnumerable<CommentNode> Comments { get; }

        /// <summary>
        /// Gets a copy of the collected declarations.
        /// </summary>
        IEnumerable<Declaration> AllDeclarations { get; }

        IEnumerable<KeyValuePair<VBComponent, IParseTree>> ParseTrees { get; }

        void SetModuleState(VBComponent component, ParserState state, SyntaxErrorException parserError = null);
        ParserState GetModuleState(VBComponent component);
        void SetModuleComments(VBComponent component, IEnumerable<CommentNode> comments);

        /// <summary>
        /// Adds the specified <see cref="Declaration"/> to the collection (replaces existing).
        /// </summary>
        void AddDeclaration(Declaration declaration);

        void ClearDeclarations(VBComponent component);
        void AddTokenStream(VBComponent component, ITokenStream stream);
        void AddParseTree(VBComponent component, IParseTree parseTree);
        TokenStreamRewriter GetRewriter(VBComponent component);

        /// <summary>
        /// Ensures parser state accounts for built-in declarations.
        /// This method has no effect if built-in declarations have already been loaded.
        /// </summary>
        void AddBuiltInDeclarations(IHostApplication hostApplication);
    }
}