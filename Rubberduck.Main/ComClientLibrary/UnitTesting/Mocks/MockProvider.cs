using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;
using Moq;
using Rubberduck.Resources.Registration;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;

// ReSharper disable InconsistentNaming

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IMockProviderGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IMockProvider
    {
        [DispId(1)]
        IComMock Mock(string ProgId, [Optional] string ProjectName);

        [DispId(2)]
        MockArgumentCreator It();
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.MockProviderGuid),
        ProgId(RubberduckProgId.MockProviderProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IMockProvider))
    ]
    public class MockProvider : IMockProvider
    {
        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = true)]
        private static extern int CLSIDFromProgID(string lpszProgID, out Guid lpclsid);

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = true)]
        private static extern int LoadTypeLib(string fileName, out ITypeLib typeLib);

        private readonly MockArgumentCreator _it = new MockArgumentCreator();
        private static readonly ConcurrentDictionary<string, Type> typeCache = new ConcurrentDictionary<string, Type>();

        public IComMock Mock(string ProgId, string ProjectName = null)
        {
            var key = string.Concat(ProjectName, "::", ProgId);
            if (!typeCache.TryGetValue(key, out var classType))
            {
                // In order to mock a COM type, we must acquire a Type. However,
                // ProgId will only return the coclass, which itself is a collection
                // of interfaces, so we must take additional steps to obtain the default
                // interface rather than the class itself.
                classType = string.IsNullOrWhiteSpace(ProjectName)
                    ? Type.GetTypeFromProgID(ProgId)
                    : GetVbaType(ProgId, ProjectName);

                if (classType == null)
                {
                    throw new ArgumentOutOfRangeException(nameof(ProgId),
                        $"The supplied {ProgId} was not found. The class may not be registered.");
                }

                if (classType.Name == "__ComObject")
                {
                    if (TryGetTypeInfoFromProgId(ProgId, out var typeInfo))
                    {
                        var pUnk = Marshal.GetIUnknownForObject(typeInfo);
                        classType = Marshal.GetTypeForITypeInfo(pUnk);
                        Marshal.Release(pUnk);

                        if (classType == null)
                        {
                            throw new ArgumentOutOfRangeException(nameof(ProgId),
                                $"The supplied {ProgId} was found, but we could not acquire the required metadata on the type to mock it. The class may not support early-binding.");
                        }
                    }
                }

                typeCache.TryAdd(key, classType);
            }

            var targetType = classType.IsInterface ? classType : GetComDefaultInterface(classType);

            var closedMockType = typeof(Mock<>).MakeGenericType(targetType);
            var mock = (Mock)Activator.CreateInstance(closedMockType);

            return new ComMock(mock, targetType, classType.GetInterfaces());
        }

        private static bool TryGetTypeInfoFromProgId(string progId, out ITypeInfo typeInfo)
        {
            typeInfo = null;
            if (CLSIDFromProgID(progId, out var clsid) != 0)
            {
                return false;
            }

            if (!TryGetTypeLibFromCLSID(clsid, out var lib))
            {
                return false;
            }

            lib.GetTypeInfoOfGuid(ref clsid, out typeInfo);
            return true;
        }

        private static bool TryGetTypeLibFromCLSID(Guid clsid, out ITypeLib lib)
        {
            lib = null;

            var clsidKey = Registry.ClassesRoot.OpenSubKey($"CLSID\\{clsid.ToString("B")}");
            if (clsidKey == null)
            {
                return false;
            }

            if (!TryLoadTypeLibFromPath(TryGetTypeLibPath, clsidKey, out lib))
            {
                return true;
            }

            if (TryLoadTypeLibFromPath(TryGetInProcServerPath, clsidKey, out lib))
            {
                return true;
            }

            if (TryLoadTypeLibFromPath(TryGetLocalServerPath, clsidKey, out lib))
            {
                return true;
            }

            return false;
        }

        private delegate bool GetPathFunction(RegistryKey clsidKey, out string path);
        private static bool TryLoadTypeLibFromPath(GetPathFunction getPathFunction, RegistryKey clsidKey, out ITypeLib lib)
        {
            lib = null;
            if (!getPathFunction(clsidKey, out var path))
            {
                return false;
            }

            if (LoadTypeLib(path, out lib) == 0)
            {
                return true;
            }

            var file = Path.GetFileName(path);
            return LoadTypeLib(file, out lib) == 0;
        }

        private static bool TryGetTypeLibPath(RegistryKey clsidKey, out string path)
        {
            path = null;

            var clsidTypeLibKey = clsidKey.OpenSubKey("TypeLib");
            if (clsidTypeLibKey == null)
            {
                return false;
            }

            var typeLibGuidValue = clsidTypeLibKey.GetValue(null);

            var typeLibKey = Registry.ClassesRoot.OpenSubKey($"TypeLib\\{typeLibGuidValue}");
            if (typeLibKey == null)
            {
                return false;
            }

            var versions = typeLibKey.GetSubKeyNames();
            var versionDictionary = new Dictionary<Version, string>();
            foreach (var version in versions)
            {
                if (Version.TryParse(version, out var v))
                {
                    versionDictionary.Add(v, version);
                }
            }

            var latestVersion = versionDictionary
                .OrderByDescending(x => x.Key.Major)
                .ThenByDescending(x => x.Key.Minor)
                .ThenByDescending(x => x.Key.Build)
                .ThenByDescending(x => x.Key.Revision)
                .FirstOrDefault().Value;

            var typeLibVersionKey = typeLibKey.OpenSubKey(latestVersion);
            if (typeLibVersionKey == null)
            {
                return false;
            }

            var flagKeys = typeLibVersionKey.GetSubKeyNames();
            var flag = flagKeys.Contains("0") ? "0" : flagKeys.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(flag))
            {
                return false;
            }

            var typeLibFlagKey = typeLibVersionKey.OpenSubKey(flag);
            var syskind = Environment.Is64BitProcess ? "win64" : "win32";

            var typeLibKindKey = typeLibFlagKey.OpenSubKey(syskind);
            if (typeLibKindKey == null)
            {
                return false;
            }

            path = typeLibKindKey.GetValue(null) as string;
            return !string.IsNullOrWhiteSpace(path);
        }

        private static bool TryGetInProcServerPath(RegistryKey clsidKey, out string path)
        {
            var procServerKey = clsidKey.OpenSubKey("InprocServer32");
            if (procServerKey != null)
            {
                path = procServerKey.GetValue(null) as string;
                return true;
            }

            path = null;
            return false;
        }

        private static bool TryGetLocalServerPath(RegistryKey clsidKey, out string path)
        {
            var localServerKey = clsidKey.OpenSubKey("LocalServer32");
            if (localServerKey != null)
            {
                path = localServerKey.GetValue(null) as string;
                return true;
            }

            path = null;
            return false;
        }

        public MockArgumentCreator It() => _it;

        private static Type GetComDefaultInterface(Type classType)
        {
            Type targetType = null;

            var pTI = Marshal.GetITypeInfoForType(classType);
            var ti = (ITypeInfo) Marshal.GetTypedObjectForIUnknown(pTI, typeof(ITypeInfo));
            ti.GetTypeAttr(out var attr);
            using (DisposalActionContainer.Create(attr, ptr => ti.ReleaseTypeAttr(ptr)))
            {
                var typeAttr = Marshal.PtrToStructure<TYPEATTR>(attr);
                if (typeAttr.typekind == TYPEKIND.TKIND_COCLASS && typeAttr.cImplTypes > 0)
                {
                    for (var i = 0; i < typeAttr.cImplTypes; i++)
                    {
                        ti.GetImplTypeFlags(i, out var implTypeFlags);

                        if ((implTypeFlags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) !=
                            IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT ||
                            (implTypeFlags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FRESTRICTED) ==
                            IMPLTYPEFLAGS.IMPLTYPEFLAG_FRESTRICTED ||
                            (implTypeFlags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) ==
                            IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE)
                        {
                            continue;
                        }

                        ti.GetRefTypeOfImplType(i, out var href);
                        ti.GetRefTypeInfo(href, out var iTI);

                        iTI.GetDocumentation(-1, out var strName, out _, out _, out _);

                        targetType = classType.GetInterface(strName, true);
                    }
                }
            }

            if (targetType == null)
            {
                // Could not find the default interface using type infos, so we'll just pick
                // whatever's the first listed and hope for the best.
                targetType = classType.GetInterfaces().FirstOrDefault() ?? classType;
            }

            return targetType;
        }

        private static Type GetVbaType(string ProgId, string ProjectName)
        {
            Type classType = null;

            if (!TryGetVbeProject(ProjectName, out var project))
            {
                return null;
            }

            var lib = TypeLibWrapper.FromVBProject(project);

            foreach (var info in lib.TypeInfos)
            {
                if (info.Name != ProgId)
                {
                    continue;
                }

                var typeInfo = (ITypeInfo) info;
                var pTypeInfo = Marshal.GetComInterfaceForObject(typeInfo, typeof(ITypeInfo));

                // TODO: find out why this crashes with NRE; the pointer seems to be valid, but
                // the exception comes from deep within the mscorlib assembly. It might not
                // be liking some of funkiness that the VBA class typeinfo generates.
                //
                // Note: Tried both TypeLibConverter class and TypeToTypeInfoMarshaler class;
                // all go into the same code path that throws NRE. 
                classType = Marshal.GetTypeForITypeInfo(pTypeInfo);
                Marshal.Release(pTypeInfo);
                break;
            }

            return classType;
        }

        private static bool TryGetVbeProject(string ProjectName, out IVBProject project)
        {
            var vbe = VbeProvider.Vbe;
            project = null;
            using (var projects = vbe.VBProjects)
            {
                foreach (var proj in projects)
                {
                    if (proj.Name != ProjectName)
                    {
                        proj.Dispose();
                        continue;
                    }

                    project = proj;
                    break;
                }
            }

            return project != null;
        }
    }
}
