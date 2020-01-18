using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Symbols;

// ReSharper disable once CheckNamespace; namespace MUST be the same as Antlr-generated VBAParser types.
namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class AnnotationContext : IAnnotatingContext
        {
            public ParserRuleContext AnnotatedContext { get; internal set; }

            public string AnnotationType => Identifier.GetName(this.annotationName().unrestrictedIdentifier());
        }

        public partial class ModuleAttributesContext : IAnnotatedContext // holds module-scoped annotations
        {
            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public int AttributeTokenIndex => this.Stop.TokenIndex + 1;

            public void Annotate(AnnotationContext annotation)
            {
                _annotations.Add(annotation);
            }

            public void AddAttributes(Attributes attributes)
            {
                foreach(var attribute in attributes)
                {
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class SubStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => block()?.Start.TokenIndex ?? Stop.TokenIndex + 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class FunctionStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext

            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }

            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => block()?.Start.TokenIndex ?? Stop.TokenIndex + 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class DeclareStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext

            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }

            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Stop.TokenIndex + 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class EventStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext

            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }

            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
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
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion
        }

        public partial class ConstSubStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class VariableSubStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion 

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class PropertyGetStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => block()?.Start.TokenIndex ?? Stop.TokenIndex + 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class PropertyLetStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => block()?.Start.TokenIndex ?? Stop.TokenIndex + 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class PropertySetStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => block()?.Start.TokenIndex ?? Stop.TokenIndex + 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class EnumerationStmtContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class EnumerationStmt_ConstantContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class UdtDeclarationContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class UdtMemberContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class IdentifierStatementLabelContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }

        public partial class SimpleNameExprContext : IIdentifierContext, IAnnotatedContext
        {
            #region IIdentifierContext
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out Interval tokenInterval);
                    return tokenInterval;
                }
            }
            #endregion

            #region IAnnotatedContext
            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

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
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
            #endregion
        }
    }
}
