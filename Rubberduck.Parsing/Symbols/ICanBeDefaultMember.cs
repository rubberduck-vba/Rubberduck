using System.Linq;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Symbols
{
    public interface ICanBeDefaultMember
    {
        string IdentifierName { get; }
        Attributes Attributes { get; }
        /// <summary>
        /// Gets an attribute value indicating whether a member is a class' default member.
        /// If this value is true, any reference to an instance of the class it's the default member of,
        /// should count as a member call to this member.
        /// </summary>
        bool IsDefaultMember { get; }       
    }

    internal static class CanBeDefaultMember
    {
        /// <summary>
        /// Provides a default implementation of ICanBeDefaultMember.IsDefaultMember.
        /// </summary>
        /// <param name="member">The member to test.</param>
        /// <returns></returns>
        public static bool IsDefaultMember(this ICanBeDefaultMember member) => member.Attributes.Any(a =>
            a.Name == $"{member.IdentifierName}.VB_UserMemId" && a.Values.Single() == "0");
    }
}
