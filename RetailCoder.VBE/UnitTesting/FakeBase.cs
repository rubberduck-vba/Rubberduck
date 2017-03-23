using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UnitTesting
{
    internal abstract class FakeBase : StubBase, IFake
    {
        #region Internal

        protected struct ReturnValueInfo
        {
            public int Invocation { get; }
            public string Parameter { get; }
            public object Argument { get; }
            public object ReturnValue { get; }

            public ReturnValueInfo(int invocation, string parameter, object argument, object returns)
            {
                Invocation = invocation;
                Parameter = parameter.ToLower();
                Argument = argument;
                ReturnValue = returns;
            }
        }

        internal object ReturnValue { get; set; }
        internal bool SuppressesCall { get; set; } = true;

        protected override void TrackUsage(string parameter, object value, string typeName)
        {
            base.TrackUsage(parameter, value, typeName);

            if (TrySetReturnValue(parameter, value) ||          // specific value, parameter, invocation
                TrySetReturnValue(parameter, value, true) ||    // specific value, parameter, any invocation
                TrySetReturnValue() ||                          // specific invocation
                TrySetReturnValue(true))                        // any invocation    
            {
                SuppressesCall = true;
                return;
            }
            SuppressesCall = false;
        }

        private bool TrySetReturnValue(string parameter, object value, bool any = false)
        {
            var returnInfo =
                ReturnValues.Where(r => r.Invocation == (any ? FakesProvider.AllInvocations : (int) InvocationCount) &&
                                        r.Argument != null &&
                                        r.Parameter.Equals(parameter.ToLower()) &&
                                        r.Argument.Equals(value)).ToList();
            if (returnInfo.Count <= 0)
            {
                return false;
            }
            ReturnValue = returnInfo.First().ReturnValue;
            return true;
        }

        private bool TrySetReturnValue(bool any = false)
        {
            var returnInfo =
                ReturnValues.Where(r => r.Invocation == (any ? FakesProvider.AllInvocations : (int) InvocationCount))
                    .ToList();

            if (returnInfo.Count <= 0)
            {
                return false;
            }
            ReturnValue = returnInfo.First().ReturnValue;
            return true;
        }

        #endregion

        #region IFake

        private static readonly List<ReturnValueInfo> ReturnValues = new List<ReturnValueInfo>();
        public virtual void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            ReturnValues.Add(new ReturnValueInfo(invocation, string.Empty, string.Empty, value));
        }

        public virtual void ReturnsWhen(string parameter, object argument, object value, int invocation = FakesProvider.AllInvocations)
        {
            ReturnValues.Add(new ReturnValueInfo(invocation, parameter, argument, value));
        }

        #endregion
    }
}
