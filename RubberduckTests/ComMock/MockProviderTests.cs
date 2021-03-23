using System;
using System.Runtime.InteropServices;
using Moq;
using NUnit.Framework;
using Rubberduck.ComClientLibrary.UnitTesting.Mocks;
using Rubberduck.Parsing.ComReflection.TypeLibReflection;
using Rubberduck.Resources.Registration;

namespace RubberduckTests.ComMock
{
    [TestFixture]
    [Category("ComMocks")]
    public class MockProviderTests
    {
        // NOTE: this class includes unit tests that deals with COM internals. To ensure 
        // the tests work, it's best to stick to only COM objects that are a part of 
        // Windows such as Scripting library (scrrun.dll)

        private const string CLSID_FileSystemObject_String = "0D43FE01-F093-11CF-8940-00A0C9054228";
        private const string IID_FileSystem3_String = "2A0B9D10-4B87-11D3-A97A-00104B365C9F";
        private const string IID_FileSystem_String = "0AB5A3D0-E5B6-11D0-ABF5-00A0C90FFFC0";

        private static Guid CLSID_FileSystemObject = new Guid(CLSID_FileSystemObject_String);
        private static Guid IID_IFileSystem3 = new Guid(IID_FileSystem3_String);
        private static Guid IID_IFileSystem = new Guid(IID_FileSystem_String);

        [Test]
        public void MockProvider_Returns_ComMocked()
        {
            var provider = new MockProvider();
            var mock = provider.Mock("Scripting.FileSystemObject");
            var obj = mock.Object;

            Assert.IsInstanceOf<ComMocked>(obj);
        }

        [Test]
        public void MockProvider_Create_Correct_Mock()
        {
            var pUnk = IntPtr.Zero;
            var pCom = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);
                Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pCom);
                Assert.AreNotEqual(IntPtr.Zero, pCom);
            }
            finally
            {
                if (pCom != IntPtr.Zero) Marshal.Release(pCom);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        public void Mocked_Implements_IDispatch()
        {
            var pUnk = IntPtr.Zero;
            var pDis = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);
                var iid = new Guid("{00020400-0000-0000-C000-000000000046}");
                var hr = Marshal.QueryInterface(pUnk, ref iid, out pDis);

                Assert.AreEqual(0, hr);
                Assert.AreNotEqual(IntPtr.Zero, pDis);
            }
            finally
            {
                if (pDis != IntPtr.Zero) Marshal.Release(pDis);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        public void Mocked_Implements_IComMocked()
        {
            var pUnk = IntPtr.Zero;
            var pDis = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);
                var iid = new Guid(RubberduckGuid.IComMockedGuid);
                var hr = Marshal.QueryInterface(pUnk, ref iid, out pDis);

                Assert.AreEqual(0, hr);
                Assert.AreNotEqual(IntPtr.Zero, pDis);
            }
            finally
            {
                if (pDis != IntPtr.Zero) Marshal.Release(pDis);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        [TestCase("Scripting.FileSystemObject", "Scripting.FileSystemObject")]
        [TestCase("Scripting.FileSystemObject", "scripting.filesystemobject")]
        [TestCase("Scripting.FileSystemObject", "Scripting.Filesystemobject")]
        [TestCase("Scripting.FileSystemObject", "SCRIPTING.FILESYSTEMOBJECT")]
        [TestCase("Scripting.FileSystemObject", "sCrIpTiNg.FiLeSyStEmObJeCt")]
        public void Mock_Returns_Same_Type(string input1, string input2)
        {
            var provider1 = new MockProvider();
            var provider2 = new MockProvider();

            var mock1 = provider1.Mock(input1);
            var mock2 = provider2.Mock(input2);

            Assert.AreEqual(mock1.Object.GetType(), mock2.Object.GetType());
        }

        [Test]
        public void Mock_NoSetup_Returns_Null()
        {
            var pUnk = IntPtr.Zero;
            var pMocked = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);
                var hr = Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pMocked);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the mocked type");
                }

                dynamic proxy = Marshal.GetObjectForIUnknown(pMocked);
                
                Assert.IsNull(proxy.GetTempName());
                Assert.IsNull(proxy.BuildPath("abc", "def"));
            }
            finally
            {
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
                if (pMocked != IntPtr.Zero) Marshal.Release(pMocked);
            }
        }

        [Test]
        public void Mock_Setup_No_Args_Returns_Specified_Value()
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var expected = "foo";
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                mock.SetupWithReturns("GetTempName", expected);
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);
                
                var hr = Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the proxy type");
                }

                dynamic mocked = Marshal.GetObjectForIUnknown(pProxy);
                Assert.AreEqual(expected, mocked.GetTempName());
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        [TestCase("abc", "def")]
        [TestCase("", "")]
        [TestCase(null, null)]
        [TestCase("abc", null)]
        [TestCase(null, "def")]
        [TestCase("abc", "")]
        [TestCase("", "def")]
        public void Mock_Setup_Args_Returns_Specified_Value(string input1, string input2)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var expected = "foobar";
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                mock.SetupWithReturns("BuildPath", expected, new object[] {provider.It.IsAny(), provider.It.IsAny()});
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var hr = Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the proxy type");
                }

                dynamic mocked = Marshal.GetObjectForIUnknown(pProxy);
                Assert.AreEqual(expected, mocked.BuildPath(input1, input2));
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        [TestCase("foobar", "abc", "def")]
        [TestCase("foobar", "", "")]
        [TestCase(null, null, null)]
        [TestCase(null, "abc", null)]
        [TestCase(null, null, "def")]
        [TestCase("foobar", "abc", "")]
        [TestCase("foobar", "", "def")]
        public void Mock_Setup_NonEmpty_Args_Returns_Specified_Value(string expected, string input1, string input2)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                mock.SetupWithReturns("BuildPath", expected, new object[] { provider.It.IsNotNull(), provider.It.IsNotNull() });
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var hr = Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the proxy type");
                }

                dynamic mocked = Marshal.GetObjectForIUnknown(pProxy);
                Assert.AreEqual(expected, mocked.BuildPath(input1, input2));
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        [TestCase("foobar" ,"abc", "def")]
        [TestCase(null, "def", "abc")]
        [TestCase(null, "", "")]
        [TestCase(null, null, null)]
        [TestCase(null, "abc", null)]
        public void Mock_Setup_Specified_Args_Returns_Specified_Value(string expected, string input1, string input2)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                mock.SetupWithReturns("BuildPath", "foobar", new object[] { provider.It.Is("abc"), provider.It.Is("def") });
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var hr = Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the proxy type");
                }

                dynamic mocked = Marshal.GetObjectForIUnknown(pProxy);
                Assert.AreEqual(expected, mocked.BuildPath(input1, input2));
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        [TestCase("foobar", "foo", "bar", new[] { "foo", "baz" }, new[] { "bar", "baz" })]
        [TestCase("foobar", "baz", "baz", new[] { "foo", "baz" }, new[] { "bar", "baz" })]
        [TestCase("foobar", "foo", "bar", new[] { "foo" }, new[] { "bar" })]
        [TestCase(null, "bar", "foo", new[] { "foo" }, new[] { "bar" })]
        [TestCase(null, "derp", "duh", new[] { "foo", "baz" }, new[] { "bar", "baz" })]
        public void Mock_Setup_Args_List_Returns_Specified_Value(string expected, string input1, string input2, string[] list1, string[] list2)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                mock.SetupWithReturns("BuildPath", "foobar", new object[] { provider.It.IsIn(list1), provider.It.IsIn(list2) });
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var hr = Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the proxy type");
                }

                dynamic mocked = Marshal.GetObjectForIUnknown(pProxy);
                Assert.AreEqual(expected, mocked.BuildPath(input1, input2));
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        [TestCase(null, "foo", "bar", new[] { "foo", "baz" }, new[] { "bar", "baz" })]
        [TestCase(null, "baz", "baz", new[] { "foo", "baz" }, new[] { "bar", "baz" })]
        [TestCase(null, "foo", "bar", new[] { "foo" }, new[] { "bar" })]
        [TestCase("foobar", "bar", "foo", new[] { "foo" }, new[] { "bar" })]
        [TestCase("foobar", "derp", "duh", new[] { "foo", "baz" }, new[] { "bar", "baz" })]
        public void Mock_Setup_Args_NotInList_Returns_Specified_Value(string expected, string input1, string input2, string[] list1, string[] list2)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                mock.SetupWithReturns("BuildPath", "foobar", new object[] { provider.It.IsNotIn(list1), provider.It.IsNotIn(list2) });
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var hr = Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the proxy type");
                }

                dynamic mocked = Marshal.GetObjectForIUnknown(pProxy);
                Assert.AreEqual(expected, mocked.BuildPath(input1, input2));
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        [TestCase("foobar", "a", "d", SetupArgumentRange.Inclusive, "a", "c", "d", "f")]
        [TestCase("foobar", "b", "e", SetupArgumentRange.Inclusive, "a", "c", "d", "f")]
        [TestCase("foobar", "c", "f", SetupArgumentRange.Inclusive, "a", "c", "d", "f")]
        [TestCase("foobar", "c", "d", SetupArgumentRange.Inclusive, "a", "c", "d", "f")]
        [TestCase("foobar", "a", "f", SetupArgumentRange.Inclusive, "a", "c", "d", "f")]
        [TestCase(null, "d", "d", SetupArgumentRange.Inclusive, "a", "c", "d", "f")]
        [TestCase(null, "a", "a", SetupArgumentRange.Inclusive, "a", "c", "d", "f")]
        [TestCase("foobar", "b", "e", SetupArgumentRange.Exclusive, "a", "c", "d", "f")]
        [TestCase(null, "a", "e", SetupArgumentRange.Exclusive, "a", "c", "d", "f")]
        [TestCase(null, "c", "e", SetupArgumentRange.Exclusive, "a", "c", "d", "f")]
        [TestCase(null, "b", "d", SetupArgumentRange.Exclusive, "a", "c", "d", "f")]
        [TestCase(null, "b", "f", SetupArgumentRange.Exclusive, "a", "c", "d", "f")]
        public void Mock_Setup_Args_Range_Returns_Specified_Value(string expected, string input1, string input2, SetupArgumentRange type, string start1, string end1, string start2, string end2)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                mock.SetupWithReturns("BuildPath", "foobar", new object[] { provider.It.IsInRange(start1, end1, type), provider.It.IsInRange(start2, end2, type)});
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var hr = Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the proxy type");
                }

                dynamic mocked = Marshal.GetObjectForIUnknown(pProxy);
                Assert.AreEqual(expected, mocked.BuildPath(input1, input2));
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        public void Mock_FileSystemObject_Via_ComMocked()
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                if (!CachedTypeService.Instance.TryGetCachedType("Scripting.FileSystemObject", out var targetType))
                {
                    throw new InvalidOperationException("Unable to locate the ProgId `Scripting.FileSystemObject`");
                }
                var closedMockType = typeof(Mock<>).MakeGenericType(targetType);
                var mock = (Mock)Activator.CreateInstance(closedMockType);
                var mockedProvider = new Mock<IMockProviderInternal>();
                var comMock = new Rubberduck.ComClientLibrary.UnitTesting.Mocks.ComMock(mockedProvider.Object, string.Empty, "Scripting.FileSystemObject", mock, targetType, targetType.GetInterfaces());
                var comMocked = new ComMocked(comMock, targetType.GetInterfaces());
                var obj = comMocked;

                pUnk = Marshal.GetIUnknownForObject(obj);
                var hr = Marshal.QueryInterface(pUnk, ref IID_IFileSystem3, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed");
                }

                dynamic proxy = Marshal.GetObjectForIUnknown(pProxy);

                Assert.IsInstanceOf(targetType, proxy);
                foreach (var face in targetType.GetInterfaces())
                {
                    Assert.IsInstanceOf(face, proxy);
                }
                Assert.AreNotEqual(pUnk, pProxy);
                Assert.AreNotSame(obj, proxy);
                Assert.IsInstanceOf<ComMocked>(obj);
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        [TestCase(IID_FileSystem3_String, true, "abc")]
        [TestCase(IID_FileSystem_String, true, "abc")]
        [TestCase(IID_FileSystem3_String, false, "def")]
        [TestCase(IID_FileSystem_String, false, "def")]
        [TestCase(IID_FileSystem3_String, false, "")]
        [TestCase(IID_FileSystem_String, false, "")]
        public void Mock_Setup_Property_Specified_Args_Returns_Specified_Object(string IID, bool expected, string input)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mockFso = provider.Mock("Scripting.FileSystemObject");
                var mockDrives = mockFso.SetupChildMock("Drives");
                var mockDrive = mockDrives.SetupChildMock("Item", provider.It.Is("abc"));
                mockDrive.SetupWithReturns("Path", "foobar");
                var obj = mockFso.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var iid = new Guid(IID);
                var hr = Marshal.QueryInterface(pUnk, ref iid, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the proxy type");
                }

                dynamic mocked = Marshal.GetObjectForIUnknown(pProxy);
                Assert.AreEqual(expected, mocked.Drives[input] != null);
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        [Test]
        [TestCase(IID_FileSystem3_String, "foobar", "abc")]
        [TestCase(IID_FileSystem_String, "foobar", "abc")]
        public void Mock_Setup_Property_Specified_Args_Returns_Specified_Value(string IID, string expected, string input)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;
            
            try
            {
                var provider = new MockProvider();
                var mockFso = provider.Mock("Scripting.FileSystemObject");
                var mockDrives = mockFso.SetupChildMock("Drives");
                var mockDrive = mockDrives.SetupChildMock("Item",  provider.It.Is("abc"));
                mockDrive.SetupWithReturns("Path", "foobar");
                var obj = mockFso.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var iid = new Guid(IID);
                var hr = Marshal.QueryInterface(pUnk, ref iid, out pProxy);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the proxy type");
                }

                dynamic mocked = Marshal.GetObjectForIUnknown(pProxy);
                Assert.AreEqual(expected, mocked.Drives[input].Path);
            }
            finally
            {
                if (pProxy != IntPtr.Zero) Marshal.Release(pProxy);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        /* Commented to remove the PIA reference to Scripting library, but keeping code in one day they fix type equivalence?
        [Test]
        public void Type_From_ITypeInfo_Are_Equivalent()
        {
            var other = typeof(FileSystemObject);
            var service = new TypeLibQueryService();
            if (service.TryGetTypeInfoFromProgId("Scripting.FileSystemObject", out var type))
            {
                Assert.IsTrue(type.IsEquivalentTo(typeof(FileSystemObject)));
            }
        }
        */
    }
}
