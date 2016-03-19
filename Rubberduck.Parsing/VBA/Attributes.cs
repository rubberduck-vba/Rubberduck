using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA
{
    /// <summary>
    /// A dictionary storing values for a given attribute.
    /// </summary>
    /// <remarks>
    /// Dictionary key is the attribute name/identifier.
    /// </remarks>
    public class Attributes : Dictionary<string, IEnumerable<string>>
    {
        /// <summary>
        /// Specifies that a member is the type's default member.
        /// Only one member in the type is allowed to have this attribute.
        /// </summary>
        /// <param name="identifierName"></param>
        public void AddDefaultMemberAttribute(string identifierName)
        {
            Add(identifierName + ".VB_UserMemId", new[] {"0"});
        }

        public void AddHiddenMemberAttribute(string identifierName)
        {
            Add(identifierName + ".VB_UserMemId", new[] {"40"});
        }

        public void AddEnumeratorMemberAttribute(string identifierName)
        {
            Add(identifierName + ".VB_UserMemId", new[] {"-4"});
        }

        public void AddMemberDescriptionAttribute(string identifierName, string description)
        {
            Add(identifierName + ".VB_Description", new[] {description});
        }

        public void AddPredeclaredIdTypeAttribute()
        {
            Add("VB_PredeclaredId", new[] {"True"});
        }
    }
}