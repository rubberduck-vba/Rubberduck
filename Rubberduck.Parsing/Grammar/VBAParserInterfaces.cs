using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class AnnotationContext : IChildContext
        {
            public ParserRuleContext ParentContext { get; internal set; }
        }

        public partial class SubStmtContext : IIdentifierContext, IMemberContext
        {
            #region IIdentifierContext

            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IMemberContext
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public void Annotate(AnnotationContext annotation)
            {
                _annotations.Add(annotation);
            }

            public void AddAttributes(Attributes attributes)
            {
                foreach (var attribute in attributes)
                {
                    Attributes.Add(attribute.Key, attribute.Value);
                }
            }
            #endregion
        }

        public partial class FunctionStmtContext : IIdentifierContext, IMemberContext
        {
            #region IIdentifierContext

            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }

            #endregion

            #region IMemberContext
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public void Annotate(AnnotationContext annotation)
            {
                _annotations.Add(annotation);
            }

            public void AddAttributes(Attributes attributes)
            {
                foreach(var attribute in attributes)
                {
                    Attributes.Add(attribute.Key, attribute.Value);
                }
            }
            #endregion
        }

        public partial class EventStmtContext : IIdentifierContext, IMemberContext
        {
            #region IIdentifierContext

            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }

            #endregion

            #region IMemberContext
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public void Annotate(AnnotationContext annotation)
            {
                _annotations.Add(annotation);
            }

            public void AddAttributes(Attributes attributes)
            {
                foreach(var attribute in attributes)
                {
                    Attributes.Add(attribute.Key, attribute.Value);
                }
            }
            #endregion
        }


        public partial class ArgContext : IIdentifierContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion
        }

        public partial class ConstSubStmtContext : IIdentifierContext, IMemberContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IMemberContext
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public void Annotate(AnnotationContext annotation)
            {
                _annotations.Add(annotation);
            }

            public void AddAttributes(Attributes attributes)
            {
                foreach(var attribute in attributes)
                {
                    Attributes.Add(attribute.Key, attribute.Value);
                }
            }
            #endregion
        }

        public partial class VariableSubStmtContext : IIdentifierContext, IMemberContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion 

            #region IMemberContext
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public void Annotate(AnnotationContext annotation)
            {
                _annotations.Add(annotation);
            }

            public void AddAttributes(Attributes attributes)
            {
                foreach(var attribute in attributes)
                {
                    Attributes.Add(attribute.Key, attribute.Value);
                }
            }
            #endregion
        }

        public partial class PropertyGetStmtContext : IIdentifierContext, IMemberContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IMemberContext
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public void Annotate(AnnotationContext annotation)
            {
                _annotations.Add(annotation);
            }

            public void AddAttributes(Attributes attributes)
            {
                foreach(var attribute in attributes)
                {
                    Attributes.Add(attribute.Key, attribute.Value);
                }
            }
            #endregion
        }

        public partial class PropertyLetStmtContext : IIdentifierContext, IMemberContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IMemberContext
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public void Annotate(AnnotationContext annotation)
            {
                _annotations.Add(annotation);
            }

            public void AddAttributes(Attributes attributes)
            {
                foreach(var attribute in attributes)
                {
                    Attributes.Add(attribute.Key, attribute.Value);
                }
            }
            #endregion
        }

        public partial class PropertySetStmtContext : IIdentifierContext, IMemberContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IMemberContext
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public void Annotate(AnnotationContext annotation)
            {
                _annotations.Add(annotation);
            }

            public void AddAttributes(Attributes attributes)
            {
                foreach(var attribute in attributes)
                {
                    Attributes.Add(attribute.Key, attribute.Value);
                }
            }
            #endregion
        }
    }
}
