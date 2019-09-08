using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

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

        public AttributeNode(VBAParser.AttributeStmtContext context)
        {
            Context = context;
            Name = context?.attributeName().GetText() ?? string.Empty;
            _values = context?.attributeValue().Select(a => a.GetText()).ToList() ?? new List<string>();
        }

        public AttributeNode(string attributeName, IEnumerable<string> values)
        {
            Name = attributeName;
            _values = values.ToList();
        }

        public VBAParser.AttributeStmtContext Context { get; }

        public string Name { get; }

        //The order of the values matters for some attributes, e.g. VB_Ext_Key.
        public IReadOnlyList<string> Values => _values.AsReadOnly();

        public bool HasValue(string value)
        {
            return _values.Any(item => item.Equals(value, StringComparison.OrdinalIgnoreCase));
        }

        public bool Equals(AttributeNode other)
        {
            return other != null 
                   && string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase) 
                   && _values.SequenceEqual(other.Values);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as AttributeNode);
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(_values.Concat(new []{Name.ToUpperInvariant()}));
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
        public static bool IsDefaultAttribute(ComponentType componentType, string attributeName, IReadOnlyList<string> attributeValues)
        {
            if (!ComponentsWithDefaultAttributes.Contains(componentType))
            {
                return false;
            }

            switch (attributeName)
            {
                case "VB_Name":
                    return true;
                case "VB_GlobalNameSpace":
                    return (componentType == ComponentType.ClassModule || componentType == ComponentType.UserForm)
                           && attributeValues[0].Equals(Tokens.False);
                case "VB_Exposed":
                    return (componentType == ComponentType.ClassModule || componentType == ComponentType.UserForm)
                           && attributeValues[0].Equals(Tokens.False);
                case "VB_Creatable":
                    return (componentType == ComponentType.ClassModule || componentType == ComponentType.UserForm)
                           && attributeValues[0].Equals(Tokens.False);
                case "VB_PredeclaredId":
                    return (componentType == ComponentType.ClassModule && attributeValues[0].Equals(Tokens.False))
                           || (componentType == ComponentType.UserForm && attributeValues[0].Equals(Tokens.True));
                default:
                    return false;
            }
        }

        private static readonly ICollection<ComponentType> ComponentsWithDefaultAttributes = new HashSet<ComponentType>
        {
            ComponentType.StandardModule,
            ComponentType.ClassModule,
            ComponentType.UserForm
        };

        public static string AttributeBaseName(string attributeName, string memberName)
        {
            return attributeName.StartsWith($"{memberName}.")
                ? attributeName.Substring(memberName.Length + 1)
                : attributeName;
        }

        public static string MemberAttributeName(string attributeBaseName, string memberName)
        {
            return $"{memberName}.{attributeBaseName}";
        }

        public bool HasAttributeFor(IParseTreeAnnotation annotation, string memberName = null)
        {
            return AttributeNodesFor(annotation, memberName).Any();
        }

        public IEnumerable<AttributeNode> AttributeNodesFor(IParseTreeAnnotation annotationInstance, string memberName = null)
        {
            if (!(annotationInstance.Annotation is IAttributeAnnotation annotation))
            {
                return Enumerable.Empty<AttributeNode>();
            }
            var attribute = annotation.Attribute(annotationInstance);

            var attributeName = memberName != null
                ? MemberAttributeName(attribute, memberName)
                : attribute;
            //VB_Ext_Key annotation depend on the defined key for identity.
            if (attribute.Equals("VB_Ext_Key", StringComparison.OrdinalIgnoreCase))
            {
                return this.Where(a => a.Name.Equals(attributeName, StringComparison.OrdinalIgnoreCase)
                                     && a.Values[0] == annotation.AttributeValues(annotationInstance)[0]);
            }

            return this.Where(a => a.Name.Equals(attributeName, StringComparison.OrdinalIgnoreCase));
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
            Add(new AttributeNode(identifierName + ".VB_MemberFlags", new[] {"40"}));
        }

        public bool HasHiddenMemberAttribute(string identifierName, out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.HasValue("40")
                && a.Name.Equals($"{identifierName}.VB_MemberFlags", StringComparison.OrdinalIgnoreCase));
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

        public bool HasEvaluateMemberAttribute(string identifierName, out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.HasValue("-5")
                                                  && a.Name.Equals($"{identifierName}.VB_UserMemId", StringComparison.OrdinalIgnoreCase));
            return attribute != null;
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
            Add(new AttributeNode("VB_PredeclaredId", new[] {Tokens.True}));
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
            attribute = this.SingleOrDefault(a => a.Name.Equals("VB_PredeclaredId", StringComparison.OrdinalIgnoreCase) && a.HasValue(Tokens.True));
            return attribute != null;
        }

        /// <summary>
        /// Corresponds to TYPEFLAG_FAPPOBJECT?
        /// </summary>
        public void AddGlobalClassAttribute()
        {
            Add(new AttributeNode("VB_GlobalNamespace", new[] {Tokens.True}));
        }

        public bool HasGlobalAttribute(out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.Name.Equals("VB_GlobalNamespace", StringComparison.OrdinalIgnoreCase) && a.HasValue(Tokens.True));
            return attribute != null;
        }

        public void AddExposedClassAttribute()
        {
            Add(new AttributeNode("VB_Exposed", new[] { Tokens.True }));
        }

        public bool HasExposedAttribute(out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.Name.Equals("VB_Exposed", StringComparison.OrdinalIgnoreCase) && a.HasValue(Tokens.True));
            return attribute != null;
        }

        /// <summary>
        /// Corresponds to *not* having the TYPEFLAG_FNONEXTENSIBLE flag set (which is the default for VBA).
        /// </summary>
        public void AddExtensibleClassAttribute()
        {
            Add(new AttributeNode("VB_Customizable", new[] { Tokens.True }));
        }

        public bool HasExtensibleAttribute(out AttributeNode attribute)
        {
            attribute = this.SingleOrDefault(a => a.Name.Equals("VB_Customizable", StringComparison.OrdinalIgnoreCase) && a.HasValue(Tokens.True));
            return attribute != null;
        }

        public bool HasAttribute(string attribute)
        {
            return this.Any(attributeNode => attributeNode.Name.Equals(attribute, StringComparison.OrdinalIgnoreCase));
        }

        public IEnumerable<AttributeNode> AttributeNodes(string attribute)
        {
            return this.Where(attributeNode => attributeNode.Name.Equals(attribute, StringComparison.OrdinalIgnoreCase));
        }

        public IEnumerable<VBAParser.AttributeStmtContext> AttributeContexts(string attribute)
        {
            return this.Where(attributeNode => attributeNode.Name.Equals(attribute, StringComparison.OrdinalIgnoreCase)).Select(node => node.Context);
        }
    }
}
