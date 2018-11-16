using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Parsing.Symbols
{
    //Needed for VB6 support, although 1 and 2 are applicable to VBA.  See 5.2.4.1.1 https://msdn.microsoft.com/en-us/library/ee177292.aspx
    //and 5.2.4.1.2 https://msdn.microsoft.com/en-us/library/ee199159.aspx
    public enum Instancing
    {
        Private = 1,                    //Not exposed via COM.
        PublicNotCreatable = 2,         //TYPEFLAG_FCANCREATE not set.
        SingleUse = 3,                  //TYPEFLAGS.TYPEFLAG_FAPPOBJECT
        GlobalSingleUse = 4,            
        MultiUse = 5,
        GlobalMultiUse = 6
    }

    public class AttributeNode : IEquatable<AttributeNode>
    {
        private readonly IList<string> _values;
        private readonly HashSet<VBAParser.AttributeStmtContext> _contexts = new HashSet<VBAParser.AttributeStmtContext>();

        public AttributeNode(VBAParser.AttributeStmtContext context)
        {
            _contexts.Add(context);
            Name = context?.attributeName().GetText() ?? string.Empty;
            _values = context?.attributeValue().Select(a => a.GetText()).ToList() ?? new List<string>();
        }

        public AttributeNode(string name, IEnumerable<string> values)
        {
            Name = name;
            _values = values.ToList();
        }

        public IReadOnlyCollection<VBAParser.AttributeStmtContext> Contexts => _contexts;

        public string Name { get; }
    
        public void AddValue(string value)
        {
            _values.Add(value);
        }

        public void AddContext(VBAParser.AttributeStmtContext context)
        {
            if (context == null || !Name.Equals(context.attributeName().GetText(), StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            _contexts.Add(context);
            var values = context.attributeValue().Select(e => e.GetText().Replace("\"", string.Empty)).ToList();
            foreach (var value in values.Where(v => !HasValue(v)))
            {
                AddValue(value);
            }
        }

        public IReadOnlyCollection<string> Values => _values.AsReadOnly();

        public bool HasValue(string value)
        {
            return _values.Any(item => item.Equals(value, StringComparison.OrdinalIgnoreCase));
        }

        public bool Equals(AttributeNode other)
        {
            return other != null && string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as AttributeNode);
        }

        public override int GetHashCode()
        {
            return Name.GetHashCode();
        }
    }

    /// <summary>
    /// A dictionary storing values for a given attribute.
    /// </summary>
    /// <remarks>
    /// Dictionary key is the attribute name/identifier.
    /// </remarks>
    public class Attributes : HashSet<AttributeNode>
    {
        public bool HasAttributeFor(AnnotationType annotationType, string memberName = null)
        {
            if (!annotationType.HasFlag(AnnotationType.Attribute))
            {
                return false;
            }

            var name = "VB_" + annotationType;

            if (annotationType.HasFlag(AnnotationType.MemberAnnotation))
            {
                //Debug.Assert(memberName != null);
                return this.Any(a => a.Name.Equals($"{memberName}{name}", StringComparison.OrdinalIgnoreCase));
            }

            if (annotationType.HasFlag(AnnotationType.ModuleAnnotation))
            {
                //Debug.Assert(memberName == null);
                return this
                    .Any(a => a.Name.Equals(name, StringComparison.OrdinalIgnoreCase)
                           && a.Values.Any(v => v.Equals("True", StringComparison.OrdinalIgnoreCase)));
            }

            return false;
        }

        /// <summary>
        /// Specifies that a member is the type's default member.
        /// Only one member in the type is allowed to have this attribute.
        /// </summary>
        /// <param name="identifierName"></param>
        public void AddDefaultMemberAttribute(string identifierName)
        {
            Add(new AttributeNode(identifierName + ".VB_UserMemId", new[] {"0"}));
        }

        public bool HasDefaultMemberAttribute(string identifierName, out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.HasValue("0") 
                && a.Name.Equals($"{identifierName}.VB_UserMemId", StringComparison.OrdinalIgnoreCase));
            return attribute != null;
        }

        public bool HasDefaultMemberAttribute()
        {
            return this.Any(a => a.HasValue("0") 
                && a.Name.EndsWith(".VB_UserMemId", StringComparison.OrdinalIgnoreCase));
        }

        public void AddHiddenMemberAttribute(string identifierName)
        {
            Add(new AttributeNode(identifierName + ".VB_UserMemId", new[] {"40"}));
        }

        public bool HasHiddenMemberAttribute(string identifierName, out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.HasValue("-4") &&
                a.Name.Equals($"{identifierName}.VB_UserMemId", StringComparison.OrdinalIgnoreCase));
            return attribute != null;
        }

        public void AddEnumeratorMemberAttribute(string identifierName)
        {
            Add(new AttributeNode(identifierName + ".VB_UserMemId", new[] {"-4"}));
        }

        /// <summary>
        /// Corresponds to DISPID_EVALUATE
        /// </summary>
        /// <param name="identifierName"></param>
        public void AddEvaluateMemberAttribute(string identifierName)
        {
            Add(new AttributeNode(identifierName + ".VB_UserMemId", new[] { "-5" }));
        }

        public void AddMemberDescriptionAttribute(string identifierName, string description)
        {
            Add(new AttributeNode(identifierName + ".VB_Description", new[] {description}));
        }

        public bool HasMemberDescriptionAttribute(string identifierName, out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.Name.Equals($"{identifierName}.VB_Description", StringComparison.OrdinalIgnoreCase));
            return attribute != null;
        }

        /// <summary>
        /// Corresponds to TYPEFLAGS.TYPEFLAG_FPREDECLID
        /// </summary>
        public void AddPredeclaredIdTypeAttribute()
        {
            Add(new AttributeNode("VB_PredeclaredId", new[] {"True"}));
        }

        public AttributeNode PredeclaredIdAttribute
        {
            get
            {
                return this.SingleOrDefault(a => a.Name.Equals("VB_PredeclaredId", StringComparison.OrdinalIgnoreCase));
            }
        }

        public AttributeNode ExposedAttribute
        {
            get
            {
                return this.SingleOrDefault(a => a.Name.Equals("VB_Exposed", StringComparison.OrdinalIgnoreCase));
            }
        }

        public bool HasPredeclaredIdAttribute(out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.Name.Equals("VB_PredeclaredId", StringComparison.OrdinalIgnoreCase) && a.HasValue("True"));
            return attribute != null;
        }

        /// <summary>
        /// Corresponds to TYPEFLAG_FAPPOBJECT?
        /// </summary>
        public void AddGlobalClassAttribute()
        {
            Add(new AttributeNode("VB_GlobalNamespace", new[] {"True"}));
        }

        public bool HasGlobalAttribute(out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.Name.Equals("VB_GlobalNamespace", StringComparison.OrdinalIgnoreCase) && a.HasValue("True"));
            return attribute != null;
        }

        public void AddExposedClassAttribute()
        {
            Add(new AttributeNode("VB_Exposed", new[] { "True" }));
        }

        public bool HasExposedAttribute(out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.Name.Equals("VB_Exposed", StringComparison.OrdinalIgnoreCase) && a.HasValue("True"));
            return attribute != null;
        }

        /// <summary>
        /// Corresponds to *not* having the TYPEFLAG_FNONEXTENSIBLE flag set (which is the default for VBA).
        /// </summary>
        public void AddExtensibledClassAttribute()
        {
            Add(new AttributeNode("VB_Customizable", new[] { "True" }));
        }

        public bool HasExtensibleAttribute(out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.Name.Equals("VB_Customizable", StringComparison.OrdinalIgnoreCase) && a.HasValue("True"));
            return attribute != null;
        }
    }
}
