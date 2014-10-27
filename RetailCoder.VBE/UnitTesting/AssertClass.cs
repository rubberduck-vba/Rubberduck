using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.UnitTesting
{
    [ComVisible(true)]
    [ComDefaultInterface(typeof(IAssert))]
    public class AssertClass : IAssert
    {
        public void IsTrue(bool condition, string message = null)
        {
            if (condition)
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("IsTrue", message);
            }
        }

        public void IsFalse(bool condition, string message = null)
        {
            if (!condition)
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("IsFalse", message);
            }
        }

        public void Inconclusive(string message = null)
        {
            AssertHandler.OnAssertInconclusive(message);
        }

        public void Fail(string message = null)
        {
            AssertHandler.OnAssertFailed("Fail", message);
        }

        public void IsNothing(object value, string message = null)
        {
            if (value == null)
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("IsNothing", message);
            }
        }

        public void IsNotNothing(object value, string message = null)
        {
            if (value != null)
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("IsNotNothing", message);
            }
        }

        public void AreEqual(object value1, object value2, string message = null)
        {
            if (value1.Equals(value2))
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("AreEqual", message);
            }
        }

        public void AreNotEqual(object value1, object value2, string message = null)
        {
            if (!value1.Equals(value2))
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("AreNotEqual", message);
            }
        }

        public void AreSame(object value1, object value2, string message = null)
        {
            if (object.ReferenceEquals(value1, value2))
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("AreSame", message);
            }
        }

        public void AreNotSame(object value1, object value2, string message = null)
        {
            if (!object.ReferenceEquals(value1, value2))
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("AreNotSame", message);
            }
        }
    }
}
