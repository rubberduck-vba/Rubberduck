namespace Rubberduck.Parsing.Symbols
{
    public interface ICanBeInterfaceMember
    {
        /// <summary>
        /// Returns true if the member is part of an interface definition.
        /// </summary>
        bool IsInterfaceMember { get; }
        Accessibility Accessibility { get; }
        Declaration ParentDeclaration { get; }
        /// <summary>
        /// Returns the Declaration of the interface that this is a member of.
        /// </summary>
        ClassModuleDeclaration InterfaceDeclaration { get; }
        string IdentifierName { get; }
        DeclarationType DeclarationType { get; }
        bool IsObject { get; }
        string AsTypeName { get; }
    }

    internal static class InterfaceMemberExtensions
    {
        /// <summary>
        /// Provides a default implementation of ICanBeInterfaceMember.IsInterfaceMember
        /// </summary>
        /// <param name="member"></param>
        /// <returns></returns>
        internal static bool IsInterfaceMember(this ICanBeInterfaceMember member) =>
            (member.Accessibility == Accessibility.Public || member.Accessibility == Accessibility.Implicit) &&
            member.InterfaceDeclaration != null;

        /// <summary>
        /// Provides a default implementation of ICanBeInterfaceMember.InterfaceDeclaration
        /// </summary>
        /// <param name="member">The member to find the InterfaceDeclaration of.</param>
        /// <returns></returns>
        internal static ClassModuleDeclaration InterfaceDeclaration(this ICanBeInterfaceMember member) =>
            member.ParentDeclaration is ClassModuleDeclaration parent && parent.IsInterface ? parent : null;
    }
}
