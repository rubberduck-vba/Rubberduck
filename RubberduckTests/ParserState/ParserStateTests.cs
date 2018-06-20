using System;
using System.Threading;
using System.Threading.Tasks;
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

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_NewTask_IsQueued()
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
                Task.Run(() =>
                {
                    Thread.Sleep(50);
                    state.OnParseRequested(this);
                });
                Thread.Sleep(100);
            });

            Assert.IsTrue(wasReparsed);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_Interrupted_IsQueued()
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var wasBusy = false;
            var reparseAfterBusy = 0;
            Task result = null;

            state.StateChanged += (o, e) =>
            {
                if (e.State == ParserState.Started)
                {
                    if (result == null)
                    {
                        result = Task.Run(() =>
                        {
                            state.OnSuspendParser(this, () =>
                            {
                                wasBusy = state.Status == ParserState.Busy;
                            });
                        });
                        wasBusy = false;
                    }
                }

                if (e.State == ParserState.Ready && wasBusy)
                {
                    reparseAfterBusy++;
                }
            };

            state.OnParseRequested(this);
            while (result == null)
            {
                Thread.Sleep(1);
            }
            result.Wait();

            Assert.AreEqual(1, reparseAfterBusy);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_Interrupted_Deadlock()
        {            
            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var wasSuspended = false;
            var wasSuspensionExecuted = false;

            // The cancellation token exists primarily to prevent
            // unwanted inlining of the tasks.
            // See: https://stackoverflow.com/questions/12245935/is-task-factory-startnew-guaranteed-to-use-another-thread-than-the-calling-thr
            var source = new CancellationTokenSource();
            var token = source.Token;
            Task result2 = null;

            state.StateChanged += (o, e) =>
            {
                if (e.State == ParserState.Started)
                {
                    result2 = Task.Run(() =>
                    {
                        wasSuspensionExecuted =
                            state.OnSuspendParser(this, () => { wasSuspended = state.Status == ParserState.Busy; },
                                20);
                    }, token);
                    result2.Wait(token);
                }
            };
            var result1 = Task.Run(() =>
            {
                state.OnParseRequested(this);
            }, token);
            result1.Wait(token);
            while (result2 == null)
            {
                Thread.Sleep(1);
            }
            result2.Wait(token);
            Assert.IsFalse(wasSuspended, "wasSuspended was set to true");
            Assert.IsFalse(wasSuspensionExecuted, "wasSuspensionExecuted was set to true");
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_Interrupted_TwoRequests_IsQueued()
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var wasRunning = false;
            var wasBusy = false;
            var reparseAfterBusy = 0;
            Task result1 = null;
            Task result2 = null;
            Task suspendTask = null;

            state.StateChanged += (o, e) =>
            {
                if (e.State == ParserState.Started && !wasRunning)
                {
                    if (result1 == null)
                    {
                        result1 = Task.Run(() =>
                        {
                            wasRunning = true;
                            result2 = Task.Run(() => state.OnParseRequested(this));
                        });
                        return;
                    }
                }

                if (e.State == ParserState.Started && wasRunning)
                {
                    suspendTask = Task.Run(() =>
                    {
                        state.OnSuspendParser(this, () =>
                        {
                            wasBusy = state.Status == ParserState.Busy;
                        });
                    });
                    return;
                }

                if (e.State == ParserState.Ready && wasBusy)
                {
                    reparseAfterBusy++;
                }
            };

            state.OnParseRequested(this);
            while (result1 == null)
            {
                Thread.Sleep(1);
            }
            result1.Wait();
            while (result2 == null)
            {
                Thread.Sleep(1);
            }
            result2.Wait();
            while (suspendTask == null)
            {
                Thread.Sleep(1);
            }
            suspendTask.Wait();
            Assert.AreEqual(1, reparseAfterBusy);
        }
    }
}
