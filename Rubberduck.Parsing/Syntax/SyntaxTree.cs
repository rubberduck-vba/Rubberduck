using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Syntax
{
    public interface ISyntaxFactory
    {
        SyntaxTree Create(ISyntaxFactory factory, WeakReference<SyntaxTree> parent, TextSpan span, IReadOnlyList<SyntaxTree> children);
    }

    public abstract class SyntaxTree
    {
        private readonly ISyntaxFactory _factory;
        private readonly WeakReference<SyntaxTree> _parent;
        private readonly IReadOnlyList<SyntaxTree> _children;

        protected SyntaxTree(ISyntaxFactory factory, WeakReference<SyntaxTree> parent, TextSpan span)
            : this(factory, parent, span, null)
        {
        }

        private SyntaxTree(ISyntaxFactory factory, WeakReference<SyntaxTree> parent, TextSpan span, IReadOnlyList<SyntaxTree> children)
        {
            _factory = factory;
            _parent = parent;
            _span = span;
            _children = children;
        }

        private SyntaxTree Create(TextSpan span, IReadOnlyList<SyntaxTree> children)
        {
            return _factory.Create(_factory, _parent, span, children);
        }

        public SyntaxTree AddChild(SyntaxTree child)
        {
            var span = TextSpan.Expand(_span, child.Span);
            var children = _children != null 
                ? _children.Concat(new[] {child}).ToArray() 
                : new[] {child};

            return Create(span, children);
        }

        public SyntaxTree InsertChild(int index, SyntaxTree child)
        {
            if (_children == null) { throw new InvalidOperationException(); }
            if (index < 0 || index >= _children.Count) { throw new ArgumentOutOfRangeException(); }

            var span = TextSpan.Expand(_span, child.Span);
            var children = _children.Take(index + 1).ToList();
            children.Add(child);
            children.AddRange(_children.Except(children).Select(c => c.Shift(child.Span.Columns, child.Span.Lines)));

            return Create(span, children);
        }

        public SyntaxTree Shift(int columns, int lines = 0)
        {
            var span = new TextSpan(_span.StartLine + lines, _span.StartColumn + columns,
                                    _span.EndLine + lines, _span.EndColumn + columns);
            return Create(span, _children);
        }

        public SyntaxTree Parent
        {
            get
            {
                SyntaxTree target;
                return _parent.TryGetTarget(out target) ? target : null;
            }
        }

        private readonly TextSpan _span;
        public TextSpan Span { get { return _span; } }
    }
}