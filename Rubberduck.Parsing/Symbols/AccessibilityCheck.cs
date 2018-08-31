using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public static class AccessibilityCheck
    {
        public static bool IsAccessible(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration callee)
        {
            return callee != null 
                    && (callee.DeclarationType.HasFlag(DeclarationType.Project) 
                        || (callee.DeclarationType.HasFlag(DeclarationType.Module) && IsModuleAccessible(callingProject, callingModule, callee))
                        || (!callee.DeclarationType.HasFlag(DeclarationType.Module) && IsMemberAccessible(callingProject, callingModule, callingParent, callee)));
        }


        public static bool IsModuleAccessible(Declaration callingProject, Declaration callingModule, Declaration calleeModule)
        {
            return calleeModule != null
                    && (IsTheSameModule(callingModule, calleeModule)
                        || IsEnclosingProject(callingProject, calleeModule)
                        || (calleeModule.DeclarationType.HasFlag(DeclarationType.ProceduralModule) && !((ProceduralModuleDeclaration)calleeModule).IsPrivateModule)
                        || (!calleeModule.DeclarationType.HasFlag(DeclarationType.ProceduralModule) && ((ClassModuleDeclaration)calleeModule).IsExposed));
        }

            private static bool IsTheSameModule(Declaration callingModule, Declaration calleeModule)
            {
                return calleeModule.Equals(callingModule);
            }

            private static bool IsEnclosingProject(Declaration callingProject, Declaration calleeModule)
            {
                return calleeModule.ParentScopeDeclaration.Equals(callingProject);
            }



        public static bool IsMemberAccessible(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration calleeMember)
        {
            if (calleeMember == null)
            {
                return false;
            }    
            if (IsInstanceMemberOfModuleOrOneOfItsSupertypes(callingModule, calleeMember)
                        || IsLocalMemberOfTheCallingSubroutineOrProperty(callingParent, calleeMember))
            {
                return true;
            }
            if (!calleeMember.IsUserDefined && calleeMember.Accessibility > Accessibility.Friend)
            {
                return true;
            }
            var memberModule = Declaration.GetModuleParent(calleeMember);
            return IsModuleAccessible(callingProject, callingModule, memberModule)
                    && (calleeMember.DeclarationType.HasFlag(DeclarationType.EnumerationMember)
                        || calleeMember.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember)
                        || calleeMember.DeclarationType.HasFlag(DeclarationType.ComAlias)
                        || HasPublicScope(calleeMember)
                        || (IsEnclosingProject(callingProject, memberModule) && IsAccessibleThroughoutTheSameProject(calleeMember)));
        }

            private static bool IsInstanceMemberOfModuleOrOneOfItsSupertypes(Declaration module, Declaration member)
            {
                return IsInstanceMemberOfModule(module, member)
                       || ClassModuleDeclaration.GetSupertypes(module).Any(supertype => IsInstanceMemberOfModuleOrOneOfItsSupertypes(supertype, member));   //ClassModuleDeclaration.GetSuperTypes never returns null.
            }

                private static bool IsInstanceMemberOfModule(Declaration module, Declaration member)
                {
                    return member.ParentScopeDeclaration.Equals(module);
                }

            private static bool IsLocalMemberOfTheCallingSubroutineOrProperty(Declaration callingParent, Declaration calleeMember)
            {
                return IsSubroutineOrProperty(callingParent) && CaleeHasSameParentScopeAsCaller(callingParent, calleeMember);
            }

                private static bool IsSubroutineOrProperty(Declaration decl)
                {
                    return decl.DeclarationType.HasFlag(DeclarationType.Property)
                        || decl.DeclarationType.HasFlag(DeclarationType.Function)
                        || decl.DeclarationType.HasFlag(DeclarationType.Procedure);
                }

                private static bool CaleeHasSameParentScopeAsCaller(Declaration callingParent, Declaration calleeMember)
                {
                    return callingParent.Equals(calleeMember.ParentScopeDeclaration);
                }

            private static bool HasPublicScope(Declaration member)
            {
                return member.Accessibility == Accessibility.Public
                    || member.Accessibility == Accessibility.Global
                    || (member.Accessibility == Accessibility.Implicit && IsPublicByDefault(member));
            }

                private static bool IsPublicByDefault(Declaration member)
                { 
                    return IsSubroutineOrProperty(member) 
                            || member.DeclarationType.HasFlag(DeclarationType.Enumeration)
                            || member.DeclarationType.HasFlag(DeclarationType.UserDefinedType);
                }

            private static bool IsAccessibleThroughoutTheSameProject(Declaration member)
            {
                return HasPublicScope(member)
                    || member.Accessibility == Accessibility.Friend; 
            }
    }
}
