using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.UnitTesting
{
    [ComVisible(true)]
    public interface IAssert
    {
        void IsTrue(bool condition, string message = null);
        void IsFalse(bool condition, string message = null);
        void Inconclusive(string message = null);
        void Fail(string message = null);
        void IsNothing(object value, string message = null);
        void IsNotNothing(object value, string message = null);
        void AreEqual(object value1, object value2, string message = null);
        void AreNotEqual(object value1, object value2, string message = null);
        void AreSame(object value1, object value2, string message = null);
        void AreNotSame(object value1, object value2, string message = null);
    }

}
