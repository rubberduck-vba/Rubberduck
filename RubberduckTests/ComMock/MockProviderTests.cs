using System;
using System.Reflection;
using System.Runtime.InteropServices;
using NUnit.Framework;
using Rubberduck.ComClientLibrary.UnitTesting.Mocks;
using Rubberduck.Resources.Registration;
using Scripting;

namespace RubberduckTests.ComMock
{
    [TestFixture]
    [Category("ComMocks")]
    public class MockProviderTests
    {
        // NOTE: this class includes unit tests that deals with COM internals. To ensure 
        // the tests work, it's best to stick to only COM objects that are a part of 
        // Windows such as Scripting library (scrrun.dll)

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
                var comGuid = new Guid("2A0B9D10-4B87-11D3-A97A-00104B365C9F"); //Scripting.IFileSystem3
                Marshal.QueryInterface(pUnk, ref comGuid, out pCom);
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
                var iid = typeof(IFileSystem3).GUID;
                var hr = Marshal.QueryInterface(pUnk, ref iid, out pMocked);
                if (hr != 0)
                {
                    throw new InvalidCastException("QueryInterface failed on the mocked type");
                }

                dynamic proxy =  Marshal.GetObjectForIUnknown(pMocked);
                
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
                
                var guid = typeof(IFileSystem3).GUID;
                var hr = Marshal.QueryInterface(pUnk, ref guid, out pProxy);
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
        public void Mock_Setup_Args_Returns_Specified_Value(string input1, string input2)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var expected = "foobar";
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                mock.SetupWithReturns("BuildPath", expected, new object[] {provider.It().IsAny(), provider.It().IsAny()});
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var guid = typeof(IFileSystem3).GUID;
                var hr = Marshal.QueryInterface(pUnk, ref guid, out pProxy);
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
        [TestCase("abc", "def", "foobar")]
        [TestCase("def", "abc", null)]
        [TestCase("", "", null)]
        [TestCase(null, null, null)]
        [TestCase("abc", null, null)]
        public void Mock_Setup_Specified_Args_Returns_Specified_Value(string input1, string input2, string expected)
        {
            var pUnk = IntPtr.Zero;
            var pProxy = IntPtr.Zero;

            try
            {
                var provider = new MockProvider();
                var mock = provider.Mock("Scripting.FileSystemObject");
                mock.SetupWithReturns("BuildPath", expected, new object[] { provider.It().Is("abc"), provider.It().Is("def") });
                var obj = mock.Object;

                pUnk = Marshal.GetIUnknownForObject(obj);

                var guid = typeof(IFileSystem3).GUID;
                var hr = Marshal.QueryInterface(pUnk, ref guid, out pProxy);
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
    }
}
