using System;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.ParserStateTests
{
    [TestFixture]
    public class ParserStateTests
    {
        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_IsBusy()
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            var state = MockParser.CreateAndParse(vbe.Object);
            state.OnSuspendParser(this, () =>
            {
                Assert.IsTrue(state.Status == ParserState.Busy);
            });
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_NonReadyState_IsQueued()
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var wasSuspended = false;

            state.SetStatusAndFireStateChanged(this, ParserState.Pending, CancellationToken.None);

            state.OnSuspendParser(this, () =>
            {
                wasSuspended = true;
            });

            Assert.IsTrue(wasSuspended);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_IsQueued()
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var wasBusy = false;
            var wasReparsed = false;
            
            state.StateChanged += (o, e) =>
            {
                if (e.State == ParserState.Ready && wasBusy)
                {
                    wasReparsed = true;
                }
            };

            state.OnSuspendParser(this, () =>
            {
                wasBusy = state.Status == ParserState.Busy;
                // This is a cheap hack to avoid the multi-threading setup... Lo and behold the laziness of me
                // Please don't do this in production.
                state.OnParseRequested(this);
                Assert.IsTrue(state.Status == ParserState.Busy);
            });
            
            Assert.IsTrue(wasReparsed);
        }
    }
}
