using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.UI;

namespace Rubberduck.UnitTesting
{
    internal struct UsageInfo
    {        
        public string Parameter { get; }
        public object Value { get; }
        public string TypeName { get; }
        public uint Invocation { get; }

        public bool IsMissing
        {
            get
            {
                // the hooked procedures recieve missing parameters as VT_ERROR
                var variant = Value as ComVariant;
                return variant == null || variant.VariantType == VarEnum.VT_ERROR;
            }
        }

        public UsageInfo(string parameter, object value, string typeName, uint invocation)
        {            
            Parameter = parameter.ToLower();
            Value = value;
            TypeName = typeName.ToLower();
            Invocation = invocation;
        }
    }

    // Do NOT throw from inside this class unless absolutely necessary.  Any errors should be trapped if possible and either
    // raised as VBA errors through its ErrObject or relayed back as an inconclusive or failed test result as appropriate. If
    // the test is attempting to verify that an error is thrown (i.e Assert.Succeed in an error handler) this could lead to
    // incorrect test results.
    internal class Verifier : IVerify
    {
        private readonly List<UsageInfo> _usages = new List<UsageInfo>();

        #region Internal

        protected bool Asserted { get; set; }

        internal void SuppressAsserts()
        {
            Asserted = true;
        }

        protected uint InvocationCount
        {
            get { return !_usages.Any() ? 0 : _usages.Max(u => u.Invocation); }
        }

        internal void AddUsage(string parameter, object value, string typeName, uint invocation)
        {

            _usages.Add(new UsageInfo(parameter, value, typeName, invocation));
        }

        private UsageInfo? GetUsageOrAssert(string parameter, int invocation, string message = "", [CallerMemberName] string methodName = "")
        {
            if (invocation > InvocationCount || invocation < 1)
            {
                // ReSharper disable once ExplicitCallerInfoArgument
                AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_VerifyNoInvocationFormat, parameter, _usages.Count, message), methodName);
                Asserted = true;
                return null;
            }

            var uses = _usages.Where(u => u.Parameter.Equals(parameter.ToLower()) && u.Invocation == invocation).ToArray();
            if (uses.Length != 1)
            {
                AssertHandler.OnAssertInconclusive(RubberduckUI.Assert_VerifyInternalErrorMessage);
                Asserted = true;
                return null;
            }
            return uses[0];
        }

        private bool IsEasterEgg(object value)
        {
            if (value.GetType() == typeof(AssertClass))
            {
                AssertHandler.OnAssertInconclusive(RubberduckUI.Assert_EasterEggAssertClassPassed);
                Asserted = true;
                return true;
            }

            if (value.GetType() == typeof(IVerify))
            {
                AssertHandler.OnAssertInconclusive(RubberduckUI.Assert_EasterEggIVerifyPassed);
                Asserted = true;
                return true;
            }

            if (value.GetType() == typeof(IFake))
            {
                AssertHandler.OnAssertInconclusive(RubberduckUI.Assert_EasterEggIFakePassed);
                Asserted = true;
                return true;
            }
            return false;
        }

        #endregion

        #region IVerify

        public void AtLeast(int invocations, string message = "")
        {
            if (Asserted || InvocationCount >= invocations)
            {
                return;
            }
            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, invocations, InvocationCount, message));
            Asserted = true;
        }

        public void AtLeastOnce(string message = "")
        {
            if (Asserted || InvocationCount > 0)
            {
                return;
            }
            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, 1, 0, message));
            Asserted = true;
        }

        public void AtMost(int invocations, string message = "")
        {
            if (Asserted || InvocationCount <= invocations)
            {
                return;
            }
            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, invocations, InvocationCount, message));
            Asserted = true;
        }

        public void AtMostOnce(string message = "")
        {
            if (Asserted || InvocationCount > 1)
            {
                return;
            }
            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, 1, InvocationCount, message));
            Asserted = true;
        }

        public void Between(int minimum, int maximum, string message = "")
        {
            if (Asserted || InvocationCount >= minimum && InvocationCount <= maximum)
            {
                return;
            }
            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, $"{minimum} - {maximum}", InvocationCount, message));
            Asserted = true;
        }

        public void Exactly(int invocations, string message = "")
        {
            if (Asserted || InvocationCount != invocations)
            {
                return;
            }
            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, invocations, InvocationCount, message));
            Asserted = true;
        }

        public void Never(string message = "")
        {
            if (Asserted || InvocationCount > 0)
            {
                return;
            }
            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, 0, InvocationCount, message));
            Asserted = true;
        }

        public void Once(string message = "")
        {
            if (Asserted || InvocationCount == 1)
            {
                return;
            }
            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, 1, InvocationCount, message));
            Asserted = true;
        }

        public void Parameter(string parameter, object value, int invocation = 1, string message = "")
        {
            var usage = GetUsageOrAssert(parameter, invocation, message);
            if (Asserted || !usage.HasValue || IsEasterEgg(value) || usage.Value.Value.Equals(value))
            {
                return;
            }

            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, value, usage.Value.Value, message));
            Asserted = true;
        }

        public void ParameterInRange(string parameter, double minimum, double maximum, int invocation = 1, string message = "")
        {
            var usage = GetUsageOrAssert(parameter, invocation, message);
            if (Asserted || !usage.HasValue)
            {
                return;
            }

            if (usage.Value.IsMissing)
            {
                AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_VerifyParameterNotPassed, parameter, invocation, message));
                Asserted = true;
                return;
            }

            var underTest = usage.Value.Value is ComVariant ? ((ComVariant)(usage.Value.Value)).Value : usage.Value.Value;
            if (!(underTest is double))
            {
                AssertHandler.OnAssertInconclusive(string.Format(RubberduckUI.Assert_VerifyParameterNonNumeric, parameter, invocation, message));
                Asserted = true;
                return;
            }

            // passing case.
            if ((double)underTest >= minimum && (double)underTest <= maximum)
            {
                return;
            }

            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, $"{minimum} - {maximum}", underTest, message));
            Asserted = true;
        }

        public void ParameterIsPassed(string parameter, int invocation = 1, string message = "")
        {
            var usage = GetUsageOrAssert(parameter, invocation, message);
            if (Asserted || !usage.HasValue || !usage.Value.IsMissing)
            {
                return;
            }

            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_VerifyParameterNotPassed, parameter, invocation, message));
            Asserted = true;
        }

        public void ParameterIsType(string parameter, string typeName, int invocation = 1, string message = "")
        {
            var usage = GetUsageOrAssert(parameter, invocation, message);
            if (Asserted || !usage.HasValue)
            {
                return;
            }

            //TODO
            AssertHandler.OnAssertInconclusive(RubberduckUI.Assert_NotImplemented);
            //if (usage.Value.TypeName.ToLower().Equals(typeName.ToLower()))
            //{
            //    return;
            //}

            //AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, typeName, usage.Value.TypeName, message));
            //Asserted = true;
        }

        #endregion
    }
}
