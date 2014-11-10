using RetailCoderVBE.Reflection.VBA.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA.Grammar
{
    [Flags]
    internal enum SyntaxType
    {
        Syntax = 0,
        /// <summary>
        /// Indicates that this syntax produces child nodes.
        /// </summary>
        HasChildNodes = 1,
        /// <summary>
        /// Indicates that this syntax isn't part of the language's general grammar, 
        /// e.g. 
        /// </summary>
        IsChildNodeSyntax = 2
    }

    internal abstract class SyntaxBase : ISyntax
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="nodeFactory">
        /// A factory that creates a specific implementation of the <see cref="SyntaxTreeNode"/> abstract class.
        /// </param>
        protected SyntaxBase(SyntaxType syntaxType = SyntaxType.Syntax)
        {
            _syntaxType = syntaxType;
        }

        protected abstract bool MatchesSyntax(string instruction, out Match match);
        protected abstract SyntaxTreeNode CreateNode(string scope, Match match, string instruction, string comment);

        protected virtual string Scope(string publicScope, string localScope, Match match)
        {
            var publicScopeKeywords = new[] { 
                                                ReservedKeywords.Public, 
                                                ReservedKeywords.Global 
                                            };

            return publicScopeKeywords.Contains(match.Groups[0].Value)
                                        ? publicScope
                                        : localScope;
        }

        private readonly SyntaxType _syntaxType;
        public SyntaxType Type { get { return _syntaxType; } }

        public SyntaxTreeNode ToNode(string publicScope, string localScope, string instruction)
        {
            Match match;
            if (!MatchesSyntax(instruction, out match))
            {
                return null;
            }

            var scope = Scope(publicScope, localScope, match);
            var comment = FindComment(instruction);
         
            return CreateNode(scope, match, instruction, comment);
        }

        protected virtual string FindComment(string instruction)
        {
            var comment = string.Empty;
            int commentStart;
            if (instruction.HasComment(out commentStart))
            {
                comment = instruction.Substring(commentStart);
            }

            return comment;
        }

        public bool IsMatch(string publicScope, string localScope, string instruction, out SyntaxTreeNode node)
        {
            node = ToNode(publicScope, localScope, instruction);
            return node != null;
        }


        public bool IsChildNodeSyntax
        {
            get { return _syntaxType.HasFlag(SyntaxType.IsChildNodeSyntax); }
        }
    }
}
