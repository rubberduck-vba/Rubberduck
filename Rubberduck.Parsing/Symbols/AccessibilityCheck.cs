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

            private static bool IsValidAccessibility(Declaration moduleOrMember)
            {
                return moduleOrMember != null
                       && (moduleOrMember.Accessibility == Accessibility.Global
                           || moduleOrMember.Accessibility == Accessibility.Public
                           || moduleOrMember.Accessibility == Accessibility.Friend
                           || moduleOrMember.Accessibility == Accessibility.Implicit);
            }


        public static bool IsMemberAccessible(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration calleeMember)
        {
            if (calleeMember == null)
            {
                return false;
            }    
            else if (IsEnclosingModuleOfInstanceMember(callingModule, calleeMember) || (IsSubroutineOrProperty(callingParent) && CaleeHasSameParentAsCaller(callingParent, calleeMember)))
            {
                return true;
            }
            var memberModule = Declaration.GetModuleParent(calleeMember);
            if (IsModuleAccessible(callingProject, callingModule, memberModule))
            {
                if (calleeMember.DeclarationType.HasFlag(DeclarationType.EnumerationMember) || calleeMember.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember))
                {
                    return IsValidAccessibility(calleeMember.ParentDeclaration);
                }
                else
                {
                    return IsValidAccessibility(calleeMember);
                }
            }
            return false;
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

            private static bool IsSubroutineOrProperty(Declaration decl)
            {
                return decl.DeclarationType.HasFlag(DeclarationType.Property)
                    || decl.DeclarationType.HasFlag(DeclarationType.Function)
                    || decl.DeclarationType.HasFlag(DeclarationType.Procedure);
            }

            private static bool CaleeHasSameParentAsCaller(Declaration callingParent, Declaration calleeMember)
            {
                return callingParent.Equals(calleeMember.ParentScopeDeclaration);
            }
    }
}
