using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Moq;
using Rubberduck.Parsing.ComReflection.TypeLibReflection;
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
        [Description("Creates a new mock for the specified interface.")]
        IComMock Mock(string ProgId, [Optional] string Project);

        [DispId(2)]
        [Description("Gets an object that creates argument placeholders for an expression.")]
        SetupArgumentCreator It { get; }

        [DispId(3)]
        [Description("Gets an object that specifies how many times a verifiable invocation should occur.")]
        ITimes Times { get; }
    }

    [ComVisible(false)]
    internal interface IMockProviderInternal : IMockProvider
    {
        IComMock MockChildObject(IComMock ParentObject, Type childType);
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.MockProviderGuid),
        ProgId(RubberduckProgId.MockProviderProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IMockProvider))
    ]
    public class MockProvider : IMockProviderInternal
    {
        private static readonly ICachedTypeService TypeCacheService = CachedTypeService.Instance;

        public MockProvider()
        {
            It = new SetupArgumentCreator();
            Times = new Times();
        }

        public IComMock Mock(string ProgId, string Project = null)
        {
            // If already cached, we must re-use the type to work around the 
            // broken type equivalence.
            if (TypeCacheService.TryGetCachedType(Project, ProgId, out var classType))
            {
                return CreateComMock(Project, ProgId, classType);
            }

            // In order to mock a COM type, we must acquire a Type. However,
            // ProgId will only return the coclass, which itself is a collection
            // of interfaces, so we must take additional steps to obtain the default
            // interface rather than the class itself.
            classType = string.IsNullOrWhiteSpace(Project)
                ? Type.GetTypeFromProgID(ProgId)
                : GetVbaType(ProgId, Project);

            if (classType == null)
            {
                throw new ArgumentOutOfRangeException(nameof(ProgId),
                    $"The supplied {ProgId} was not found. The class may not be registered or could not be located with the available metadata.");
            }

            return CreateComMock(Project, ProgId, classType);
        }

        public IComMock MockChildObject(IComMock ParentObject, Type childType)
        {
            var project = ParentObject.Project;
            var progId = childType.FullName;
            childType = TypeCacheService.TryGetCachedTypeFromEquivalentType(project, progId, childType);
            return CreateComMock(project, progId, childType);
        }

        private ComMock CreateComMock(string project, string progId, Type classType)
        {
            var targetType = classType.IsInterface ? classType : GetComDefaultInterface(classType);

            var closedMockType = typeof(Mock<>).MakeGenericType(targetType);
            var mock = (Mock)Activator.CreateInstance(closedMockType);

            // Ensure that the mock implements all the interfaces to cover the case where
            // no setup is performed on the given interface and to ensure that mock can 
            // be cast successfully.
            var asGenericMemberInfo = closedMockType.GetMethod("As");
            System.Diagnostics.Debug.Assert(asGenericMemberInfo != null);

            var supportedTypes = classType.GetInterfaces();
            foreach (var type in supportedTypes)
            {
                var asMemberInfo = asGenericMemberInfo.MakeGenericMethod(type);
                asMemberInfo.Invoke(mock, null);
            }

            return new ComMock(this, project, progId, mock, targetType, supportedTypes);
        }

        public SetupArgumentCreator It { get; }
        public ITimes Times { get; }

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

        private static Type GetVbaType(string progId, string projectName)
        {
            Type classType = null;

            if (!TryGetVbeProject(projectName, out var project))
            {
                return null;
            }

            var provider = new TypeLibWrapperProviderLite();
            var lib = provider.TypeLibWrapperFromProject(project);

            foreach (var typeInfo in lib.TypeInfos)
            {
                if (typeInfo.Name != progId)
                {
                    continue;
                }

                if (TypeCacheService.TryGetCachedType(typeInfo, projectName, out classType))
                {
                    break;
                }
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
