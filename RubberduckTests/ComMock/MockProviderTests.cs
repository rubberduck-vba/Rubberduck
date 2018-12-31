using System;
using System.Runtime.InteropServices;
using NUnit.Framework;
using Rubberduck.ComClientLibrary.UnitTesting.Mocks;

namespace RubberduckTests.ComMock
{
    [TestFixture]
    public class MockProviderTests
    {
        [Test]
        [Category("ComMocks")]
        public void MockProvider_Returns_ComMocked()
        {
            var provider = new MockProvider();
            var mock = provider.Mock("Excel.Application");
            var obj = mock.Object;

            Assert.IsInstanceOf<ComMocked>(obj);
        }

        [Test]
        [Category("ComMocks")]
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
    }
}
