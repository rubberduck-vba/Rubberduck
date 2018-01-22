using System;
using System.Runtime.InteropServices;
using System.Text;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibsAbstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    // for debug purposes, just reinventing the wheel here to document the major things exposed by a particular ITypeLib
    // (compatible with all ITypeLibs, not just VBE ones, but also documents the VBE specific extensions)
    // this is a throw away class, once proper integration into RD has been achieved.
    public class TypeLibDocumenter
    {
        StringBuilder _document = new StringBuilder();

        public override string ToString() => _document.ToString();

        private void AppendLine(string value = "")
            => _document.Append(value + "\r\n");

        private void AppendLineButRemoveEmbeddedNullChars(string value)
            => AppendLine(value.Replace("\0", string.Empty));

        public void AddTypeLib(ComTypes.ITypeLib typeLib)
        {
            AppendLine();
            AppendLine("================================================================================");
            AppendLine();

            string libName;
            string libString;
            int libHelp;
            string libHelpFile;
            typeLib.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out libName, out libString, out libHelp, out libHelpFile);

            libName = libName ?? "[VBA.Immediate.Window]";

            AppendLine("ITypeLib: " + libName);
            if (libString != null) AppendLineButRemoveEmbeddedNullChars("- Documentation: " + libString);
            if (libHelp != 0) AppendLineButRemoveEmbeddedNullChars("- HelpContext: " + libHelp);
            if (libHelpFile != null) AppendLineButRemoveEmbeddedNullChars("- HelpFile: " + libHelpFile);

            IntPtr typeLibAttributesPtr;
            typeLib.GetLibAttr(out typeLibAttributesPtr);
            var typeLibAttributes = StructHelper.ReadStructure<ComTypes.TYPELIBATTR>(typeLibAttributesPtr);
            typeLib.ReleaseTLibAttr(typeLibAttributesPtr);          // no need to keep open.  copied above

            AppendLine("- Guid: " + typeLibAttributes.guid);
            AppendLine("- Lcid: " + typeLibAttributes.lcid);
            AppendLine("- SysKind: " + typeLibAttributes.syskind);
            AppendLine("- LibFlags: " + typeLibAttributes.wLibFlags);
            AppendLine("- MajorVer: " + typeLibAttributes.wMajorVerNum);
            AppendLine("- MinorVer: " + typeLibAttributes.wMinorVerNum);

            var typeLibVBE = typeLib as TypeLibWrapper;
            if (typeLibVBE != null)
            {
                AppendLine("- HasVBEExtensions: " + typeLibVBE.HasVBEExtensions());
                if (typeLibVBE.HasVBEExtensions())
                {
                    AppendLine("- VBE Conditional Compilation Arguments: " + typeLibVBE.ConditionalCompilationArguments);
                }
            }

            int CountOfTypes = typeLib.GetTypeInfoCount();
            AppendLine("- TypeCount: " + CountOfTypes);

            for (int typeIdx = 0; typeIdx < CountOfTypes; typeIdx++)
            {
                ComTypes.ITypeInfo typeInfo;
                typeLib.GetTypeInfo(typeIdx, out typeInfo);

                AddTypeInfo(typeInfo, libName, 0);
            }
        }

        void AddTypeInfo(ComTypes.ITypeInfo typeInfo, string qualifiedName, int implementsLevel)
        {
            AppendLine();
            if (implementsLevel == 0)
            {
                AppendLine("-------------------------------------------------------------------------------");
                AppendLine();
            }
            implementsLevel++;

            IntPtr typeAttrPtr = IntPtr.Zero;
            typeInfo.GetTypeAttr(out typeAttrPtr);
            var typeInfoAttributes = StructHelper.ReadStructure<ComTypes.TYPEATTR>(typeAttrPtr);
            typeInfo.ReleaseTypeAttr(typeAttrPtr);

            string typeName = null;
            string typeString = null;
            int typeHelp = 0;
            string TypeHelpFile = null;

            typeInfo.GetDocumentation((int)TypeLibConsts.MEMBERID_NIL, out typeName, out typeString, out typeHelp, out TypeHelpFile);

            AppendLine(qualifiedName + "::" + (typeName.Replace("\0", string.Empty) ?? "[unnamed]"));
            if (typeString != null) AppendLineButRemoveEmbeddedNullChars("- Documentation: " + typeString.Replace("\0", string.Empty));
            if (typeHelp != 0) AppendLineButRemoveEmbeddedNullChars("- HelpContext: " + typeHelp);
            if (TypeHelpFile != null) AppendLineButRemoveEmbeddedNullChars("- HelpFile: " + TypeHelpFile.Replace("\0", string.Empty));
            AppendLine("- HasVBEExtensions: " + (((TypeInfoWrapper)typeInfo).HasVBEExtensions() ? "true" : "false"));   // FIXME not safe for ITypeInfos that aren't from our wrappers

            AppendLine("- Type: " + (TYPEKIND_VBE)typeInfoAttributes.typekind);
            AppendLine("- Guid: {" + typeInfoAttributes.guid + "}");

            AppendLine("- cImplTypes (implemented interfaces count): " + typeInfoAttributes.cImplTypes);
            AppendLine("- cFuncs (function count): " + typeInfoAttributes.cFuncs);
            AppendLine("- cVars (fields count): " + typeInfoAttributes.cVars);

            for (int funcIdx = 0; funcIdx < typeInfoAttributes.cFuncs; funcIdx++)
            {
                AddFunc(typeInfo, funcIdx);
            }

            for (int varIdx = 0; varIdx < typeInfoAttributes.cVars; varIdx++)
            {
                AddField(typeInfo, varIdx);
            }

            for (int implIndex = 0; implIndex < typeInfoAttributes.cImplTypes; implIndex++)
            {
                ComTypes.ITypeInfo typeInfoImpl = null;
                int href = 0;
                typeInfo.GetRefTypeOfImplType(implIndex, out href);
                typeInfo.GetRefTypeInfo(href, out typeInfoImpl);

                AppendLine("implements...");
                AddTypeInfo(typeInfoImpl, qualifiedName + "::" + typeName, implementsLevel);
            }
        }

        void AddFunc(ComTypes.ITypeInfo typeInfo, int funcIndex)
        {
            IntPtr funcDescPtr = IntPtr.Zero;
            typeInfo.GetFuncDesc(funcIndex, out funcDescPtr);
            var funcDesc = StructHelper.ReadStructure<ComTypes.FUNCDESC>(funcDescPtr);

            var names = new string[255];
            int cNames = 0;
            typeInfo.GetNames(funcDesc.memid, names, names.Length, out cNames);

            string namesInfo = names[0] + "(";

            int argIndex = 1;
            while (argIndex < cNames)
            {
                if (argIndex > 1) namesInfo += ", ";
                namesInfo += names[argIndex].Length > 0 ? names[argIndex] : "retVal";
                argIndex++;
            }

            namesInfo += ")";

            typeInfo.ReleaseFuncDesc(funcDescPtr);

            AppendLine("- member: " + namesInfo + " [id 0x" + funcDesc.memid.ToString("X") + ", " + funcDesc.invkind + "]");
        }

        void AddField(ComTypes.ITypeInfo typeInfo, int varIndex)
        {
            IntPtr varDescPtr = IntPtr.Zero;
            typeInfo.GetVarDesc(varIndex, out varDescPtr);
            var varDesc = StructHelper.ReadStructure<ComTypes.VARDESC>(varDescPtr);

            if (varDesc.memid != (int)TypeLibConsts.MEMBERID_NIL)
            {
                var names = new string[1];
                int cNames = 0;
                typeInfo.GetNames(varDesc.memid, names, names.Length, out cNames);
                AppendLine("- field: " + names[0] + " [id 0x" + varDesc.memid.ToString("X") + "]");
            }
            else
            {
                // Constants appear in the typelib with no name
                AppendLine("- constant: {unknown name}");
            }

            typeInfo.ReleaseVarDesc(varDescPtr);
        }
    }
}
