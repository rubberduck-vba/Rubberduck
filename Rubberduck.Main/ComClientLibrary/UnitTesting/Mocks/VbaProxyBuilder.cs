using System;
using System.Collections.Generic;
using System.Configuration.Assemblies;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Threading;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    internal class VbaProxyBuilder
    {
        internal Assembly BuildAssemblyFromVbaTypeLib(ComTypes.ITypeLib vbaTypeLib)
        {
            vbaTypeLib.GetDocumentation(-1, out var libName, out var libDocString, out _, out _);
            vbaTypeLib.GetLibAttr(out var ppTlibAttr);
            var attr = Marshal.PtrToStructure<ComTypes.TYPEATTR>(ppTlibAttr);
            var major = attr.wMajorVerNum;
            var minor = attr.wMinorVerNum;
            vbaTypeLib.ReleaseTLibAttr(ppTlibAttr);

            var asmFileName = libName + ".dll";

            var assemblyName = GetAssemblyNameFromTypelib(vbaTypeLib, asmFileName, null, null,
                new Version(major, minor, 0, 0), AssemblyNameFlags.None);
            var assemblyBuilder =
                CreateAssemblyForTypeLib(vbaTypeLib, asmFileName, assemblyName, true, true, false);

            // Define a dynamic module that will contain the contain the imported types.
            var strNonQualifiedAsmFileName = Path.GetFileName(asmFileName);
            var moduleBuilder = assemblyBuilder.DefineDynamicModule(strNonQualifiedAsmFileName, strNonQualifiedAsmFileName);

            var type = moduleBuilder.DefineType("");
            
            return null;
        }

        #region Helper Methods

        private static AssemblyBuilder CreateAssemblyForTypeLib(ComTypes.ITypeLib typeLib, string asmFileName, AssemblyName asmName, bool bPrimaryInteropAssembly, bool bReflectionOnly, bool bNoDefineVersionResource)
        {
            // Retrieve the current app domain.
            var currentDomain = Thread.GetDomain();

            // Retrieve the directory from the assembly file name.
            string dir = null;
            if (asmFileName != null)
            {
                dir = Path.GetDirectoryName(asmFileName);
                if (string.IsNullOrEmpty(dir))
                    dir = null;
            }

            AssemblyBuilderAccess aba;
            if (bReflectionOnly)
            {
                aba = AssemblyBuilderAccess.ReflectionOnly;
            }
            else
            {
                aba = AssemblyBuilderAccess.RunAndSave;
            }

            // Create the dynamic assembly itself.
            AssemblyBuilder asmBldr;

            var assemblyAttributes = new List<CustomAttributeBuilder>();
#if !FEATURE_CORECLR
            // mscorlib.dll must specify the security rules that assemblies it emits are to use, since by
            // default all assemblies will follow security rule set level 2, and we want to make that an
            // explicit decision.
            var securityRulesCtor = typeof(SecurityRulesAttribute).GetConstructor(new Type[] { typeof(SecurityRuleSet) });
            var securityRulesAttribute =
                new CustomAttributeBuilder(securityRulesCtor, new object[] { SecurityRuleSet.Level2 });
            assemblyAttributes.Add(securityRulesAttribute);
#endif // !FEATURE_CORECLR

            asmBldr = currentDomain.DefineDynamicAssembly(asmName, aba, dir, false, assemblyAttributes);

            // Set the Guid custom attribute on the assembly.
            SetGuidAttributeOnAssembly(asmBldr, typeLib);

            // Set the imported from COM attribute on the assembly and return it.
            SetImportedFromTypeLibAttrOnAssembly(asmBldr, typeLib);

            // Set the version information on the typelib.
            if (bNoDefineVersionResource)
            {
                SetTypeLibVersionAttribute(asmBldr, typeLib);
            }
            else
            {
                SetVersionInformation(asmBldr, typeLib, asmName);
            }

            // If we are generating a PIA, then set the PIA custom attribute.
            if (bPrimaryInteropAssembly)
                SetPIAAttributeOnAssembly(asmBldr, typeLib);

            return asmBldr;
        }

        private static AssemblyName GetAssemblyNameFromTypelib(object typeLib, string asmFileName, byte[] publicKey, StrongNameKeyPair keyPair, Version asmVersion, AssemblyNameFlags asmNameFlags)
        {
            // Extract the name of the typelib.
            string strTypeLibName = null;
            var dwHelpContext = 0;
            string strHelpFile = null;
            var pTLB = (ComTypes.ITypeLib)typeLib;
            pTLB.GetDocumentation(-1, out strTypeLibName, out string strDocString, out dwHelpContext, out strHelpFile);

            // Retrieve the name to use for the assembly.
            if (asmFileName == null)
            {
                asmFileName = strTypeLibName;
            }
            else
            {
                var strFileNameNoPath = Path.GetFileName(asmFileName);
                var strExtension = Path.GetExtension(asmFileName);

                // Validate that the extension is valid.
                var bExtensionValid = ".dll".Equals(strExtension, StringComparison.OrdinalIgnoreCase);

                // If the extension is not valid then tell the user and quit.
                if (!bExtensionValid)
                    throw new ArgumentException("Invalid file extension; must end in \".dll\"");

                // The assembly cannot contain the path nor the extension.
                asmFileName = strFileNameNoPath.Substring(0, strFileNameNoPath.Length - ".dll".Length);
            }

            // If the version information was not specified, then retrieve it from the typelib.
            if (asmVersion == null)
            {
                int major;
                int minor;
                pTLB.GetLibAttr(out var ppTLibAttr);
                var TLibAttr = Marshal.PtrToStructure<ComTypes.TYPELIBATTR>(ppTLibAttr);
                asmVersion = new Version(TLibAttr.wMajorVerNum, TLibAttr.wMinorVerNum, 0, 0);
                pTLB.ReleaseTLibAttr(ppTLibAttr);
            }

            // Create the assembly name for the imported typelib's assembly.
            var AsmName = new AssemblyName(asmFileName);
            AsmName.SetPublicKey(publicKey);
            AsmName.Version = asmVersion;
            AsmName.HashAlgorithm = AssemblyHashAlgorithm.None;
            AsmName.VersionCompatibility = AssemblyVersionCompatibility.SameMachine;
            AsmName.Flags = asmNameFlags;
            AsmName.KeyPair = keyPair;

            return AsmName;
        }
        
        private static void SetGuidAttributeOnAssembly(AssemblyBuilder asmBldr, ComTypes.ITypeLib typeLib)
        {
            // Retrieve the GuidAttribute constructor.
            var aConsParams = new Type[1] { typeof(string) };
            var GuidAttrCons = typeof(GuidAttribute).GetConstructor(aConsParams);

            // Create an instance of the custom attribute builder.
            var aArgs = new object[1] { Marshal.GetTypeLibGuid(typeLib).ToString() };
            var GuidCABuilder = new CustomAttributeBuilder(GuidAttrCons, aArgs);

            // Set the GuidAttribute on the assembly builder.
            asmBldr.SetCustomAttribute(GuidCABuilder);
        }

        private static void SetImportedFromTypeLibAttrOnAssembly(AssemblyBuilder asmBldr, object typeLib)
        {
            // Retrieve the ImportedFromTypeLibAttribute constructor.
            var aConsParams = new Type[1] { typeof(string) };
            var ImpFromComAttrCons = typeof(ImportedFromTypeLibAttribute).GetConstructor(aConsParams);

            // Retrieve the name of the typelib.
            var strTypeLibName = Marshal.GetTypeLibName((ComTypes.ITypeLib)typeLib);

            // Create an instance of the custom attribute builder.
            var aArgs = new object[1] { strTypeLibName };
            var ImpFromComCABuilder = new CustomAttributeBuilder(ImpFromComAttrCons, aArgs);

            // Set the ImportedFromTypeLibAttribute on the assembly builder.
            asmBldr.SetCustomAttribute(ImpFromComCABuilder);
        }

        private static void SetTypeLibVersionAttribute(AssemblyBuilder asmBldr, object typeLib)
        {
            var aConsParams = new Type[2] { typeof(int), typeof(int) };
            var TypeLibVerCons = typeof(TypeLibVersionAttribute).GetConstructor(aConsParams);

            // Get the typelib version
            var tlb = (ComTypes.ITypeLib)(typeLib);
            tlb.GetLibAttr(out var ppTLibAttr);
            var TLibAttr = Marshal.PtrToStructure<ComTypes.TYPELIBATTR>(ppTLibAttr);
            var major = TLibAttr.wMajorVerNum;
            var minor = TLibAttr.wMinorVerNum;
            tlb.ReleaseTLibAttr(ppTLibAttr);

            // Create an instance of the custom attribute builder.
            var aArgs = new object[2] { major, minor };
            var TypeLibVerBuilder = new CustomAttributeBuilder(TypeLibVerCons, aArgs);

            // Set the attribute on the assembly builder.
            asmBldr.SetCustomAttribute(TypeLibVerBuilder);
        }

        private static void SetVersionInformation(AssemblyBuilder asmBldr, object typeLib, AssemblyName asmName)
        {
            // Extract the name of the typelib.
            string strTypeLibName = null;
            string strDocString = null;
            var dwHelpContext = 0;
            string strHelpFile = null;
            var pTLB = (ComTypes.ITypeLib)typeLib;
            pTLB.GetDocumentation(-1, out strTypeLibName, out strDocString, out dwHelpContext, out strHelpFile);

            // Generate the product name string from the named of the typelib.
            var strProductName = string.Format(CultureInfo.InvariantCulture, "{0} - Imported from VBA library", strTypeLibName);

            // Set the OS version information.
            asmBldr.DefineVersionInfoResource(strProductName, asmName.Version.ToString(), null, null, null);

            // Set the TypeLibVersion attribute
            SetTypeLibVersionAttribute(asmBldr, typeLib);
        }

        private static void SetPIAAttributeOnAssembly(AssemblyBuilder asmBldr, object typeLib)
        {
            var pAttr = IntPtr.Zero;
            ComTypes.TYPELIBATTR Attr;
            var pTLB = (ComTypes.ITypeLib)typeLib;
            var Major = 0;
            var Minor = 0;

            // Retrieve the PrimaryInteropAssemblyAttribute constructor.
            var aConsParams = new Type[2] { typeof(int), typeof(int) };
            var PIAAttrCons = typeof(PrimaryInteropAssemblyAttribute).GetConstructor(aConsParams);

            // Retrieve the major and minor version from the typelib.
            try
            {
                pTLB.GetLibAttr(out pAttr);
                Attr = (ComTypes.TYPELIBATTR)Marshal.PtrToStructure(pAttr, typeof(TYPELIBATTR));
                Major = Attr.wMajorVerNum;
                Minor = Attr.wMinorVerNum;
            }
            finally
            {
                // Release the typelib attributes.
                if (pAttr != IntPtr.Zero)
                    pTLB.ReleaseTLibAttr(pAttr);
            }

            // Create an instance of the custom attribute builder.
            var aArgs = new object[2] { Major, Minor };
            var PIACABuilder = new CustomAttributeBuilder(PIAAttrCons, aArgs);

            // Set the PrimaryInteropAssemblyAttribute on the assembly builder.
            asmBldr.SetCustomAttribute(PIACABuilder);
        }

        #endregion
    }
}
