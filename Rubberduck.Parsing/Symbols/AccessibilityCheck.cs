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
            if (callee.DeclarationType.HasFlag(DeclarationType.Module))
            {
                return IsModuleAccessible(callingProject, callingModule, callee);
            }
            return IsMemberAccessible(callingProject, callingModule, callingParent, callee);
        }

        public static bool IsModuleAccessible(Declaration callingProject, Declaration callingModule, Declaration calleeModule)
        {
            bool validAccessibility = IsValidAccessibility(calleeModule);
            bool enclosingModule = callingModule.Equals(calleeModule);
            if (enclosingModule)
            {
                return true;
            }
            bool sameProject = callingModule.ParentScopeDeclaration.Equals(calleeModule.ParentScopeDeclaration);
            if (sameProject)
            {
                return validAccessibility;
            }
            if (calleeModule.DeclarationType.HasFlag(DeclarationType.ProceduralModule))
            {
                bool isPrivate = ((ProceduralModuleDeclaration)calleeModule).IsPrivateModule;
                return validAccessibility && !isPrivate;
            }
            else
            {
                bool isExposed = calleeModule != null && ((ClassModuleDeclaration)calleeModule).IsExposed;
                return validAccessibility && isExposed;
            }
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
            if (IsEnclosingModule(callingModule, calleeMember))
            {
                return true;
            }
            var callerIsSubroutineOrProperty = callingParent.DeclarationType.HasFlag(DeclarationType.Property)
                || callingParent.DeclarationType.HasFlag(DeclarationType.Function)
                || callingParent.DeclarationType.HasFlag(DeclarationType.Procedure);
            var calleeHasSameParent = callingParent.Equals(callingParent.ParentScopeDeclaration);
            if (callerIsSubroutineOrProperty && calleeHasSameParent)
            {
                return calleeHasSameParent;
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

        private static bool IsEnclosingModule(Declaration callingModule, Declaration calleeMember)
        {
            if (callingModule.Equals(calleeMember.ParentScopeDeclaration))
            {
                return true;
            }
            foreach (var supertype in ClassModuleDeclaration.GetSupertypes(callingModule))
            {
                if (IsEnclosingModule(supertype, calleeMember))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
