namespace Rubberduck.Parsing.Symbols
{
    public static class AccessibilityCheck
    {
        public static bool IsAccessible(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration callee)
        {
            if (callee.DeclarationType.HasFlag(DeclarationType.Project))
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
            if (IsEnclosingModuleOfModule(callingModule, calleeModule))
            {
                return true;
            }
            else if (IsInTheSameProject(callingModule, calleeModule))
            {
                return IsValidAccessibility(calleeModule);
            }
            else if (calleeModule.DeclarationType.HasFlag(DeclarationType.ProceduralModule))
            {
                bool isPrivate = ((ProceduralModuleDeclaration)calleeModule).IsPrivateModule;
                return  !isPrivate && IsValidAccessibility(calleeModule);
            }
            else
            {
                bool isExposed = calleeModule != null && ((ClassModuleDeclaration)calleeModule).IsExposed;
                return isExposed && IsValidAccessibility(calleeModule);
            }
        }

            private static bool IsEnclosingModuleOfModule(Declaration callingModule, Declaration calleeModule)
            {
                return callingModule.Equals(calleeModule);
            }

            private static bool IsInTheSameProject(Declaration callingModule, Declaration calleeModule)
            {
                return callingModule.ParentScopeDeclaration.Equals(calleeModule.ParentScopeDeclaration);
            }


        public static bool IsValidAccessibility(Declaration moduleOrMember)
        {
            return moduleOrMember != null
                   && (moduleOrMember.Accessibility == Accessibility.Global
                       || moduleOrMember.Accessibility == Accessibility.Public
                       || moduleOrMember.Accessibility == Accessibility.Friend
                       || moduleOrMember.Accessibility == Accessibility.Implicit);
        }

        public static bool IsMemberAccessible(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration calleeMember)
        {
            if (IsEnclosingModuleOfInstanceMember(callingModule, calleeMember) || (CallerIsSubroutineOrProperty(callingParent) && CaleeHasSameParentAsCaller(callingParent, calleeMember)))
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

            private static bool CallerIsSubroutineOrProperty(Declaration callingParent)
            {
                return callingParent.DeclarationType.HasFlag(DeclarationType.Property)
                    || callingParent.DeclarationType.HasFlag(DeclarationType.Function)
                    || callingParent.DeclarationType.HasFlag(DeclarationType.Procedure);
            }

            private static bool CaleeHasSameParentAsCaller(Declaration callingParent, Declaration calleeMember)
            {
                return callingParent.Equals(calleeMember.ParentScopeDeclaration);
            }
    }
}
