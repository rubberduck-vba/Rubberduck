using System.Runtime.InteropServices;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Reflection.VBA.Grammar
{
    [ComVisible(false)]
    public abstract class SyntaxBase : ISyntax
    {
        /// <summary>
        /// 
        /// </summary>
        protected SyntaxBase(SyntaxType syntaxType = SyntaxType.Syntax)
        {
            _syntaxType = syntaxType;
        }

        protected abstract bool MatchesSyntax(string instruction, out Match match);
        protected abstract SyntaxTreeNode CreateNode(Instruction instruction, string scope, Match match);

        protected virtual string Scope(string publicScope, string localScope, Match match)
        {
            var publicScopeKeywords = new[] { 
                                                ReservedKeywords.Public, 
                                                ReservedKeywords.Global 
                                            };

            return publicScopeKeywords.Contains(match.Value.Split(' ')[0] + ' ')
                                        ? publicScope
                                        : localScope;
        }

        private readonly SyntaxType _syntaxType;
        public SyntaxType Type { get { return _syntaxType; } }

        public SyntaxTreeNode ToNode(string publicScope, string localScope, Instruction instruction)
        {
            Match match;
            if (!MatchesSyntax(instruction.Value.Trim(), out match))
            {
                return null;
            }

            var scope = Scope(publicScope, localScope, match);
         
            return CreateNode(instruction, scope, match);
        }

        public bool IsMatch(string publicScope, string localScope, Instruction instruction, out SyntaxTreeNode node)
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
