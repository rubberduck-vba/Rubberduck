using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Utility
{
    /// <summary>
    /// All the Document() methods in one extension class for convenience
    /// </summary>
    /// <remarks>
    /// TODO: could make this more friendly by outputting in IDL format
    /// </remarks>
    internal static class TypeInfoDocumentationExtensions
    {
        public static void Document(this ITypeLibWrapper _this, StringLineBuilder output)
        {
            output.AppendLine();
            output.AppendLine("================================================================================");
            output.AppendLine();

            output.AppendLine("ITypeLib: " + _this.Name);
            output.AppendLineNoNullChars("- Documentation: " + _this.DocString);
            output.AppendLineNoNullChars("- HelpContext: " + _this.HelpContext);
            output.AppendLineNoNullChars("- HelpFile: " + _this.HelpFile);
            output.AppendLine("- Guid: " + _this.Attributes.guid);
            output.AppendLine("- Lcid: " + _this.Attributes.lcid);
            output.AppendLine("- SysKind: " + _this.Attributes.syskind);
            output.AppendLine("- LibFlags: " + _this.Attributes.wLibFlags);
            output.AppendLine("- MajorVer: " + _this.Attributes.wMajorVerNum);
            output.AppendLine("- MinorVer: " + _this.Attributes.wMinorVerNum);
            output.AppendLine("- HasVBEExtensions: " + _this.HasVBEExtensions);
            if (_this.HasVBEExtensions)
            {
                output.AppendLine("- VBE Conditional Compilation Arguments: " + _this.VBEExtensions.ConditionalCompilationArgumentsRaw);
                foreach (var reference in _this.VBEExtensions.VBEReferences)
                {
                    reference.Document(output);
                }
            }

            output.AppendLine("- TypeCount: " + _this.TypesCount);

            foreach (var typeInfo in _this.TypeInfos)
            {
                using (typeInfo)
                {
                    typeInfo.Document(output, _this.Name, 0);
                }
            }
        }

        public static void Document(this ITypeInfoWrapper _this, StringLineBuilder output, string qualifiedName, int implementsLevel)
        {
            output.AppendLine();
            if (implementsLevel == 0)
            {
                output.AppendLine("-------------------------------------------------------------------------------");
                output.AppendLine();
            }
            implementsLevel++;

            qualifiedName += "::" + (_this.Name ?? "[unnamed]");
            output.AppendLineNoNullChars(qualifiedName);
            output.AppendLineNoNullChars("- FullName: " + _this.ContainerName + "." + _this.Name);
            output.AppendLineNoNullChars("- Documentation: " + _this.DocString);
            output.AppendLineNoNullChars("- HelpContext: " + _this.HelpContext);
            output.AppendLineNoNullChars("- HelpFile: " + _this.HelpFile);

            output.AppendLine("- HasVBEExtensions: " + _this.HasVBEExtensions);
            if (_this.HasVBEExtensions) output.AppendLine("- HasModuleScopeCompilationErrors: " + _this.HasModuleScopeCompilationErrors);

            output.AppendLine("- Type: " + _this.TypeKind);
            output.AppendLine("- Guid: {" + _this.GUID + "}");

            output.AppendLine("- cImplTypes (implemented interfaces count): " + _this.ImplementedInterfaces.Count);
            output.AppendLine("- cFuncs (function count): " + _this.Funcs.Count);
            output.AppendLine("- cVars (fields count): " + _this.Vars.Count);

            foreach (var func in _this.Funcs)
            {
                using (func)
                {
                    func.Document(output);
                }
            }
            foreach (var variable in _this.Vars)
            {
                using (variable)
                {
                    variable.Document(output);
                }
            }
            foreach (var typeInfoImpl in _this.ImplementedInterfaces)
            {
                using (typeInfoImpl)
                {
                    output.AppendLine("implements...");
                    typeInfoImpl.Document(output, qualifiedName, implementsLevel);
                }
            }
        }

        public static void Document(this ITypeInfoVariable _this, StringLineBuilder output)
        {
            output.AppendLine("- field: " + _this.Name + " [id 0x" + _this.MemberID.ToString("X") + ", flags " + _this.MemberFlags.ToString() + "]");
        }

        public static void Document(this ITypeInfoFunction _this, StringLineBuilder output)
        {
            var namesInfo = _this.Name + "(";

            var namesArray = _this.NamesArray;
            var namesCount = _this.NamesArrayCount;

            var argIndex = 1;
            while (argIndex < namesCount)
            {
                if (argIndex > 1) namesInfo += ", ";
                namesInfo += namesArray[argIndex].Length > 0 ? namesArray[argIndex] : "retVal";
                argIndex++;
            }

            namesInfo += ")";

            output.AppendLine("- member: " + namesInfo + " [id 0x" + _this.MemberID.ToString("X") + ", " + _this.InvokeKind + ", flags " + _this.MemberFlags.ToString() + "]");
        }

        public static void Document(this ITypeLibReference _this, StringLineBuilder output)
        {
            output.AppendLine("- VBE Reference: " + _this.Name + " [path: " + _this.Path + ", majorVersion: " + _this.MajorVersion +
                                ", minorVersion: " + _this.MinorVersion + ", guid: " + _this.GUID + ", lcid: " + _this.LCID + "]");
        }
    }
}
