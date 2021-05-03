using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
        private static readonly IEnumerable<ParserState> AllowedRunStates = new [] {ParserState.Ready};

        [SetUp]
        public void SetUp()
        {
            // Replace pop-up assert trace listener with one that simply logs a message.
            Debug.Listeners.Clear();
            Debug.Listeners.Add(new ConsoleTraceListener());
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_IsBusy()
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                state.OnSuspendParser(this, AllowedRunStates, () =>
                {
                    Assert.IsTrue(state.Status == ParserState.Busy);
                });
            }
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_NonReadyState_IsQueued()
        {
            var wasSuspended = false;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var allowedRunStates = new[] {ParserState.Pending, ParserState.Ready};

                state.SetStatusAndFireStateChanged(this, ParserState.Pending, CancellationToken.None);

                state.OnSuspendParser(this, allowedRunStates, () =>
                {
                    wasSuspended = true;
                });
            }
            Assert.IsTrue(wasSuspended);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_IsQueued()
        {
            var wasReparsed = false;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var wasBusy = false;

                state.StateChanged += (o, e) =>
                {
                    if (e.State == ParserState.Ready && wasBusy)
                    {
                        wasReparsed = true;
                    }
                };

                state.OnSuspendParser(this, AllowedRunStates, () =>
                {
                    wasBusy = state.Status == ParserState.Busy;
                    // This is a cheap hack to avoid the multi-threading setup... Lo and behold the laziness of me
                    // Please don't do this in production.
                    state.OnParseRequested(this);
                    Assert.IsTrue(state.Status == ParserState.Busy);
                });
            }
            Assert.IsTrue(wasReparsed);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_Exception_Suspending_Inside_Parse()
        {
            var result = SuspensionOutcome.Pending;
            var wasRun = false;
            var wasSuspended = false;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                state.StateChanged += (o, e) =>
                {
                    if (e.State == ParserState.Ready && !wasRun)
                    {
                        wasRun = true;
                        result = state.OnSuspendParser(this, AllowedRunStates, () =>
                        {
                            // Cheap hack to run in same thread. Should not be done in production
                            wasSuspended = true;
                        }).Outcome;
                    }
                };
                state.OnParseRequested(this);
            }
            Assert.IsFalse(wasSuspended);
            Assert.AreEqual(SuspensionOutcome.ReadLockAlreadyHeld, result);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_Exception()
        {
            SuspensionResult result;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                result = state.OnSuspendParser(this, AllowedRunStates, () => throw new NullReferenceException());
            }

            Assert.IsNotNull(result.EncounteredException);
            Assert.AreEqual(typeof(NullReferenceException), result.EncounteredException.GetType());
            Assert.AreEqual(SuspensionOutcome.UnexpectedError, result.Outcome);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_Canceled()
        {
            SuspensionResult result;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                result = state.OnSuspendParser(this, AllowedRunStates, () => throw new OperationCanceledException());
            }

            Assert.IsNotNull(result.EncounteredException);
            Assert.AreEqual(typeof(OperationCanceledException), result.EncounteredException.GetType());
            Assert.AreEqual(SuspensionOutcome.Canceled, result.Outcome);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_IncompatibleState()
        {
            var result = SuspensionOutcome.Pending;
            var wasSuspended = false;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                result = state.OnSuspendParser(this, new []{ParserState.Pending}, () => throw new OperationCanceledException()).Outcome;
            }
            Assert.IsFalse(wasSuspended);
            Assert.AreEqual(SuspensionOutcome.IncompatibleState, result);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_NewTask_IsQueued()
        {

            var wasReparsed = false;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var wasBusy = false;

                state.StateChanged += (o, e) =>
                {
                    if (e.State == ParserState.Ready && wasBusy)
                    {
                        wasReparsed = true;
                    }
                };

                state.OnSuspendParser(this, AllowedRunStates, () =>
                {
                    wasBusy = state.Status == ParserState.Busy;
                    Task.Run(() =>
                    {
                        Thread.Sleep(50);
                        state.OnParseRequested(this);
                    });
                    Thread.Sleep(100);
                });
            }
            Assert.IsTrue(wasReparsed);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_Interrupted_IsQueued()
        {
            var reparseAfterBusy = 0;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var wasBusy = false;
                Task result = null;

                state.StateChanged += (o, e) =>
                {
                    if (e.State == ParserState.Started)
                    {
                        if (result == null)
                        {
                            result = Task.Run(() =>
                            {
                                state.OnSuspendParser(this, AllowedRunStates, () =>
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
            }

            Assert.AreEqual(1, reparseAfterBusy);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_Interrupted_Deadlock()
        {
            var wasSuspended = false;
            var result = SuspensionOutcome.Pending;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
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
                            result =
                                state.OnSuspendParser(this, AllowedRunStates,
                                    () => { wasSuspended = state.Status == ParserState.Busy; },
                                    20)
                                    .Outcome;
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
            }
            Assert.IsFalse(wasSuspended, "wasSuspended was set to true");
            Assert.AreEqual(SuspensionOutcome.TimedOut, result);
        }

        [Test]
        [Category("ParserState")]
        public void Test_RPS_SuspendParser_Interrupted_TwoRequests_IsQueued()
        {
            var reparseAfterBusy = 0;

            var vbe = MockVbeBuilder.BuildFromSingleModule("", ComponentType.StandardModule, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var wasRunning = false;
                var wasBusy = false;
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
                            state.OnSuspendParser(this, AllowedRunStates, () =>
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
            }
            Assert.AreEqual(1, reparseAfterBusy);
        }
    }
}
