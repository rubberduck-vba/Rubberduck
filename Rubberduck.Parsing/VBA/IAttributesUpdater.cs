using System.Collections.Generic;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public interface IAttributesUpdater
    {
        void AddAttribute(IRewriteSession rewriteSession, Declaration declaration, string attribute, IReadOnlyList<string> values);
        void RemoveAttribute(IRewriteSession rewriteSession, Declaration declaration, string attribute, IReadOnlyList<string> values = null);
        void UpdateAttribute(IRewriteSession rewriteSession, Declaration declaration, string attribute, IReadOnlyList<string> newValues, IReadOnlyList<string> oldValues = null);
        void AddOrUpdateAttribute(IRewriteSession rewriteSession, Declaration declaration, string attribute, IReadOnlyList<string> values);
    }
}