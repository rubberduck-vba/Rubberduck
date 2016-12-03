namespace Rubberduck.Parsing.Symbols
{
    public static class AccessibilityCheck
    {
        public static bool IsAccessible(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration callee)
        {
            if (callee == null)
            {
                return false;
            }
            else if (callee.DeclarationType.HasFlag(DeclarationType.Project))
            {
                return true;
            }
            else if (callee.DeclarationType.HasFlag(DeclarationType.Module))
            {
                return IsModuleAccessible(callingProject, callingModule, callee);
            }
            else
            {
                return IsMemberAccessible(callingProject, callingModule, callingParent, callee);
            }
        }


        public static bool IsModuleAccessible(Declaration callingProject, Declaration callingModule, Declaration calleeModule)
        {
            if (calleeModule == null)
            {
                return false;
            }    
            else if (IsTheSameModule(callingModule, calleeModule) || IsEnclosingProject(callingProject, calleeModule))
            {
                return true;
            }
            else if (calleeModule.DeclarationType.HasFlag(DeclarationType.ProceduralModule))
            {
                bool isPrivate = ((ProceduralModuleDeclaration)calleeModule).IsPrivateModule;
                return !isPrivate;
            }
            else
            {
                bool isExposed = ((ClassModuleDeclaration)calleeModule).IsExposed;
                return isExposed;
            }
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
            else if (IsEnclosingModuleOfInstanceMember(callingModule, calleeMember))
            {
                return true;
            }
            else if (IsLocalMemberOfTheCallingSubroutineOrProperty(callingParent, calleeMember))
            {
                return true;
            }
            var memberModule = Declaration.GetModuleParent(calleeMember);
            if (IsModuleAccessible(callingProject, callingModule, memberModule))
            {
                if (calleeMember.DeclarationType.HasFlag(DeclarationType.EnumerationMember) || calleeMember.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember)) 
                {
                    return true;
                }
                else if (IsEnclosingProject(callingProject, memberModule) && IsAccessibleThroughoutTheSameProject(calleeMember))
                {
                    return true;
                }
                else
                {
                    return HasPublicScope(calleeMember);
                }
            }
            else
            {
                return false;
            }
        }

            private static bool IsEnclosingModuleOfInstanceMember(Declaration callingModule, Declaration calleeMember)
            {
                if (callingModule.Equals(calleeMember.ParentScopeDeclaration))
                {
                    return true;
                }
                foreach (var supertype in ClassModuleDeclaration.GetSupertypes(callingModule))
                {
                    if (IsEnclosingModuleOfInstanceMember(supertype, calleeMember))
                    {
                        return true;
                    }
                }
                return false;
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
