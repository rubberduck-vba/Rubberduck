using Antlr4.Runtime;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Grammar;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using Rubberduck.Parsing.PreProcessing;

namespace RubberduckTests.PreProcessing
{
    [TestClass]
    public class VBAPreprocessorVisitorTests
    {
        private CultureInfo _cultureInfo;

        [TestInitialize]
        public void TestInitialize()
        {
            _cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
        }

        [TestCleanup]
        public void TestCleanup()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = _cultureInfo;
        }

        [TestMethod]
        public void TestName()
        {
            string code = @"
#Const a = 5
#Const b = a
#Const c = doesNotExist
#Const d& = 1
#Const e = d%
#Const f = [a]
";
            var result = Preprocess(code);
            Assert.AreEqual(result.Item1.Get("b"), result.Item1.Get("a"));
            Assert.AreEqual(EmptyValue.Value, result.Item1.Get("c"));
            Assert.AreEqual(1m, result.Item1.Get("d").AsDecimal);
            Assert.AreEqual(1m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(5m, result.Item1.Get("f").AsDecimal);
        }

        [TestMethod]
        public void TestMinusUnaryOperator()
        {
            string code = @"
#Const a = -5
#Const b = -#1/1/4000#
#Const c = -#1/1/2016#
#Const d = -True
#Const e = -False
#Const f = Nothing
#Const g = -""-5""
#Const h = -Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(-5m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(-767011m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(new DateTime(1783, 12, 28), result.Item1.Get("c").AsDate);
            Assert.AreEqual(1m, result.Item1.Get("d").AsDecimal);
            Assert.AreEqual(0m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("f"));
            Assert.AreEqual(5m, result.Item1.Get("g").AsDecimal);
            Assert.AreEqual(0m, result.Item1.Get("h").AsDecimal);
        }

        [TestMethod]
        public void TestNotUnaryOperator()
        {
            string code = @"
#Const a = Not 23.5
#Const b = Not #1/1/1900#
#Const c = Not False
#Const d = Nothing
#Const e = Not ""1""
#Const f = Not Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(-25m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(-3m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
            Assert.AreEqual(null, result.Item1.Get("d"));
            Assert.AreEqual(-2m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(-1m, result.Item1.Get("f").AsDecimal);
        }

        [TestMethod]
        public void TestPlusOperator()
        {
            string code = @"
#Const a = 1542.242 + 2
#Const b = #1/1/2351# + #3/6/1847#
#Const c = True + False + True
#Const d = Nothing
#Const e = ""5"" + ""4""
#Const f = Empty + Empty
#Const g = Empty + ""a""
#Const h = ""a"" + Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(1544.242m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(new DateTime(2298, 3, 7), result.Item1.Get("b").AsDate);
            Assert.AreEqual(-2m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("d"));
            Assert.AreEqual("54", result.Item1.Get("e").AsString);
            Assert.AreEqual(0m, result.Item1.Get("f").AsDecimal);
            Assert.AreEqual("a", result.Item1.Get("g").AsString);
            Assert.AreEqual("a", result.Item1.Get("h").AsString);
        }

        [TestMethod]
        public void TestMinusOperator()
        {
            string code = @"
#Const a = 10 - 8
#Const b = #1/1/2351# - #3/6/1847#
#Const c = False - True - True
#Const d = Nothing
#Const e = ""3"" - ""1""
#Const f = #1/1/2351# - 2
#Const g = #1/1/2400# - #1/1/1800# - #1/1/1800#
";
            var result = Preprocess(code);
            Assert.AreEqual(2m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(184018m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(2m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("d"));
            Assert.AreEqual(2m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(new DateTime(2350, 12, 30), result.Item1.Get("f").AsDate);
            Assert.AreEqual(new DateTime(2599, 12, 27), result.Item1.Get("g").AsDate);
        }

        [TestMethod]
        public void TestIntFunction()
        {
            string code = @"
#Const a = Int(Nothing)
#Const b = Int(Empty)
#Const c = Int(5.4)
#Const d = Int(True)
#Const e = Int(#01-30-2016#)
#Const f = Int(""10.43"")
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(0m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(5m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(-1m, result.Item1.Get("d").AsDecimal);
            Assert.AreEqual(new DateTime(2016, 1, 30), result.Item1.Get("e").AsDate);
            Assert.AreEqual(10m, result.Item1.Get("f").AsDecimal);
        }

        [TestMethod]
        public void TestFixFunction()
        {
            string code = @"
#Const a = Fix(Nothing)
#Const b = Fix(Empty)
#Const c = Fix(5.4)
#Const d = Fix(True)
#Const e = Fix(#01-30-2016#)
#Const f = Fix(""10.43"")
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(0m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(5m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(-1m, result.Item1.Get("d").AsDecimal);
            Assert.AreEqual(new DateTime(2016, 1, 30), result.Item1.Get("e").AsDate);
            Assert.AreEqual(10m, result.Item1.Get("f").AsDecimal);
        }

        [TestMethod]
        public void TestAbsFunction()
        {
            string code = @"
#Const a = Abs(Nothing)
#Const b = Abs(Empty)
#Const c = Abs(-30)
#Const d = Abs(True)
#Const e = Abs(#1/20/1005#)
#Const f = Abs(""-50"")
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(0m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(30m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(1m, result.Item1.Get("d").AsDecimal);
            Assert.AreEqual(new DateTime(2794, 12, 9), result.Item1.Get("e").AsDate);
            Assert.AreEqual(50m, result.Item1.Get("f").AsDecimal);
        }

        [TestMethod]
        public void TestSgnFunction()
        {
            string code = @"
#Const a = Sgn(-5)
#Const b = Sgn(""5"")
#Const c = Sgn(False)
#Const d = Sgn(#1/1/1855#)
#Const e = Sgn(Empty)
#Const f = Sgn(Nothing)
";
            var result = Preprocess(code);
            Assert.AreEqual(-1m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(1m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(0m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(-1m, result.Item1.Get("d").AsDecimal);
            Assert.AreEqual(0m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("f"));
        }

        [TestMethod]
        public void TestLenFunction()
        {
            string code = @"
#Const a = Len(Null)
#Const b = Len(Nothing)
#Const c = Len(Empty)
#Const d = Len(""abc"")
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(null, result.Item1.Get("b"));
            Assert.AreEqual(0m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(3m, result.Item1.Get("d").AsDecimal);
        }

        [TestMethod]
        public void TestLenBFunction()
        {
            string code = @"
#Const a = LenB(Null)
#Const b = LenB(Nothing)
#Const c = LenB(Empty)
#Const d = LenB(""abc"")
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(null, result.Item1.Get("b"));
            Assert.AreEqual(0m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(6m, result.Item1.Get("d").AsDecimal);
        }

        [TestMethod]
        public void TestCBoolFunction()
        {
            string code = @"
#Const a = CBool(Null)
#Const b = CBool(Empty)
#Const c = CBool(True)
#Const d = CBool(CByte(1))
#Const e = CBool(CByte(0))
#Const f = CBool(""tRuE"")
#Const g = CBool(""fAlSe"")
#Const h = CBool(""#TRUE#"")
#Const i = CBool(""#FALSE#"")
#Const j = CBool(""1"")
#Const k = CBool(""0"")
#Const l = CBool(#30-12-1899#)
#Const m = CBool(#31-12-1899#)
#Const n = CBool(0)
#Const o = CBool(1)
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(false, result.Item1.Get("b").AsBool);
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
            Assert.AreEqual(true, result.Item1.Get("d").AsBool);
            Assert.AreEqual(false, result.Item1.Get("e").AsBool);
            Assert.AreEqual(true, result.Item1.Get("f").AsBool);
            Assert.AreEqual(false, result.Item1.Get("g").AsBool);
            Assert.AreEqual(true, result.Item1.Get("h").AsBool);
            Assert.AreEqual(false, result.Item1.Get("i").AsBool);
            Assert.AreEqual(true, result.Item1.Get("j").AsBool);
            Assert.AreEqual(false, result.Item1.Get("k").AsBool);
            Assert.AreEqual(false, result.Item1.Get("l").AsBool);
            Assert.AreEqual(true, result.Item1.Get("m").AsBool);
            Assert.AreEqual(false, result.Item1.Get("n").AsBool);
            Assert.AreEqual(true, result.Item1.Get("o").AsBool);
        }

        [TestMethod]
        public void TestCByteFunction()
        {
            string code = @"
#Const a = CByte(Null)
#Const b = CByte(Empty)
#Const c = CByte(True)
#Const d = CByte(False)
#Const e = CByte(""1"")
#Const f = CByte(""0"")
#Const g = CByte(#30-12-1899#)
#Const h = CByte(#31-12-1899#)
#Const i = CByte(0)
#Const j = CByte(1)
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(Convert.ToByte(0), result.Item1.Get("b").AsByte);
            Assert.AreEqual(Convert.ToByte(255), result.Item1.Get("c").AsByte);
            Assert.AreEqual(Convert.ToByte(0), result.Item1.Get("d").AsByte);
            Assert.AreEqual(Convert.ToByte(1), result.Item1.Get("e").AsByte);
            Assert.AreEqual(Convert.ToByte(0), result.Item1.Get("f").AsByte);
            Assert.AreEqual(Convert.ToByte(0), result.Item1.Get("g").AsByte);
            Assert.AreEqual(Convert.ToByte(1), result.Item1.Get("h").AsByte);
            Assert.AreEqual(Convert.ToByte(0), result.Item1.Get("i").AsByte);
            Assert.AreEqual(Convert.ToByte(1), result.Item1.Get("j").AsByte);
        }

        [TestMethod]
        public void TestCAnyNumberFunction()
        {
            // Same implementation for all.
            // Simple test if any implementation of them is lacking.
            string[] functionNames = new string[]
            {
                "CCUR",
                "CDBL",
                "CINT",
                "CLNG",
                "CLNGLNG",
                "CLNGPTR",
                "CSNG"
            };
            foreach (var functionName in functionNames)
            {
                AssertIntrinsicNumberFunction(functionName);
            }
        }

        [TestMethod]
        public void TestCDateFunction()
        {
            string code = @"
#Const a = CDate(Null)
#Const b = CDate(Empty)
#Const c = CDate(True)
#Const d = CDate(False)
#Const e = CDate(""1"")
#Const f = CDate(""0"")
#Const g = CDate(1)
#Const h = CDate(0)
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(VBADateConstants.EPOCH_START, result.Item1.Get("b").AsDate);
            Assert.AreEqual(new DateTime(1899, 12, 29), result.Item1.Get("c").AsDate);
            Assert.AreEqual(VBADateConstants.EPOCH_START, result.Item1.Get("d").AsDate);
            Assert.AreEqual(new DateTime(1899, 12, 31), result.Item1.Get("e").AsDate);
            Assert.AreEqual(VBADateConstants.EPOCH_START, result.Item1.Get("f").AsDate);
            Assert.AreEqual(new DateTime(1899, 12, 31), result.Item1.Get("g").AsDate);
            Assert.AreEqual(VBADateConstants.EPOCH_START, result.Item1.Get("h").AsDate);
        }

        [TestMethod]
        public void TestCStrFunction()
        {
            string code = @"
#Const a = CStr(Null)
#Const b = CStr(Empty)
#Const c = CStr(True)
#Const d = CStr(False)
#Const e = CStr(345.23)
#Const f = CStr(#30-12-1899 02:01#)
#Const g = CStr(#1/31/2016#)
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(string.Empty, result.Item1.Get("b").AsString);
            Assert.AreEqual("True", result.Item1.Get("c").AsString);
            Assert.AreEqual("False", result.Item1.Get("d").AsString);
            Assert.AreEqual(345.23.ToString(CultureInfo.InvariantCulture), result.Item1.Get("e").AsString);
            Assert.AreEqual(new DateTime(1899, 12, 30, 2, 1, 0).ToLongTimeString(), result.Item1.Get("f").AsString);
            Assert.AreEqual(new DateTime(2016, 1, 31).ToShortDateString(), result.Item1.Get("g").AsString);
        }

        [TestMethod]
        public void TestCVariantFunction()
        {
            string code = @"
#Const a = CVAR(Null)
#Const b = CVAR(Empty)
#Const c = CVAR(True)
";
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(EmptyValue.Value, result.Item1.Get("b"));
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
        }

        private void AssertIntrinsicNumberFunction(string functionName)
        {
            string code = @"
#Const a = {0}(Null)
#Const b = {0}(Empty)
#Const c = {0}(True)
#Const d = {0}(False)
#Const e = {0}(""1"")
#Const f = {0}(""0"")
#Const g = {0}(#30-12-1899#)
#Const h = {0}(#31-12-1899#)
";
            code = string.Format(code, functionName);
            var result = Preprocess(code);
            Assert.AreEqual(null, result.Item1.Get("a"));
            Assert.AreEqual(0m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(-1m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(0m, result.Item1.Get("d").AsDecimal);
            Assert.AreEqual(1m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(0m, result.Item1.Get("f").AsDecimal);
            Assert.AreEqual(0m, result.Item1.Get("g").AsDecimal);
            Assert.AreEqual(1m, result.Item1.Get("h").AsDecimal);
        }

        [TestMethod]
        public void TestLikeOperator()
        {
            string code = @"
#Const a = ""bcabdcab"" Like ""*ab*ab""
#Const b = ""bcabdcab"" Like ""*ff*""
#Const c = ""abc"" Like ""a?c""
#Const d = ""abcd"" Like ""a?c""
#Const e = ""a1c"" Like ""a#c""
#Const f = ""abc"" Like ""a#c""
#Const g = ""a"" Like ""[!b]""
#Const h = ""a"" Like ""[!a]""
#Const i = ""1"" Like ""[0-9]""
#Const j = ""a"" Like ""[0-9]""
#Const k = Empty Like """"
#Const l = Nothing Like """"
#Const m = ""]*!"" Like ""][*a[!][a[!]""
";
            var result = Preprocess(code);
            Assert.AreEqual(true, result.Item1.Get("a").AsBool);
            Assert.AreEqual(false, result.Item1.Get("b").AsBool);
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
            Assert.AreEqual(false, result.Item1.Get("d").AsBool);
            Assert.AreEqual(true, result.Item1.Get("e").AsBool);
            Assert.AreEqual(false, result.Item1.Get("f").AsBool);
            Assert.AreEqual(true, result.Item1.Get("g").AsBool);
            Assert.AreEqual(false, result.Item1.Get("h").AsBool);
            Assert.AreEqual(true, result.Item1.Get("i").AsBool);
            Assert.AreEqual(false, result.Item1.Get("j").AsBool);
            Assert.AreEqual(true, result.Item1.Get("k").AsBool);
            Assert.AreEqual(null, result.Item1.Get("l"));
            Assert.AreEqual(true, result.Item1.Get("m").AsBool);
        }

        [TestMethod]
        public void TestIsOperator()
        {
            string code = @"
#Const a = Nothing Is Nothing
#Const b = 1 Is 2
";
            var result = Preprocess(code);
            Assert.AreEqual(true, result.Item1.Get("a").AsBool);
            Assert.AreEqual(false, result.Item1.Get("b").AsBool);
        }

        [TestMethod]
        public void TestImpOperator()
        {
            string code = @"
#Const a = False Imp False
#Const b = False Imp True
#Const c = True Imp False
#Const d = True Imp True
#Const e = #12/31/1899# Imp True
#Const f = -1 Imp Null
#Const g = -2 Imp True
#Const h = Null Imp 5
#Const i = Null Imp 0
#Const j = Null Imp Null
";
            var result = Preprocess(code);
            Assert.AreEqual(true, result.Item1.Get("a").AsBool);
            Assert.AreEqual(true, result.Item1.Get("b").AsBool);
            Assert.AreEqual(false, result.Item1.Get("c").AsBool);
            Assert.AreEqual(true, result.Item1.Get("d").AsBool);
            Assert.AreEqual(-1m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("f"));
            Assert.AreEqual(-1m, result.Item1.Get("g").AsDecimal);
            Assert.AreEqual(5m, result.Item1.Get("h").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("i"));
            Assert.AreEqual(null, result.Item1.Get("j"));
        }

        [TestMethod]
        public void TestEqvOperator()
        {
            string code = @"
#Const a = False Eqv False
#Const b = False Eqv True
#Const c = True Eqv False
#Const d = True Eqv True
#Const e = True Eqv 0
#Const f = True Eqv Null
#Const g = Null Eqv True
#Const h = Null Eqv Null
";
            var result = Preprocess(code);
            Assert.AreEqual(true, result.Item1.Get("a").AsBool);
            Assert.AreEqual(false, result.Item1.Get("b").AsBool);
            Assert.AreEqual(false, result.Item1.Get("c").AsBool);
            Assert.AreEqual(true, result.Item1.Get("d").AsBool);
            Assert.AreEqual(-0m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("f"));
            Assert.AreEqual(null, result.Item1.Get("g"));
            Assert.AreEqual(null, result.Item1.Get("h"));
        }

        [TestMethod]
        public void TestXorOperator()
        {
            string code = @"
#Const a = False Xor False
#Const b = False Xor True
#Const c = True Xor False
#Const d = True Xor True
#Const e = True Xor 0
#Const f = True Xor Null
#Const g = Null Xor True
#Const h = Null Xor Null
";
            var result = Preprocess(code);
            Assert.AreEqual(false, result.Item1.Get("a").AsBool);
            Assert.AreEqual(true, result.Item1.Get("b").AsBool);
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
            Assert.AreEqual(false, result.Item1.Get("d").AsBool);
            Assert.AreEqual(-1m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("f"));
            Assert.AreEqual(null, result.Item1.Get("g"));
            Assert.AreEqual(null, result.Item1.Get("h"));
        }

        [TestMethod]
        public void TestOrOperator()
        {
            string code = @"
#Const a = False Or False
#Const b = False Or True
#Const c = True Or False
#Const d = True Or True
#Const e = True Or 0
#Const f = True Or Null
#Const g = Null Or True
#Const h = Null Or Null
";
            var result = Preprocess(code);
            Assert.AreEqual(false, result.Item1.Get("a").AsBool);
            Assert.AreEqual(true, result.Item1.Get("b").AsBool);
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
            Assert.AreEqual(true, result.Item1.Get("d").AsBool);
            Assert.AreEqual(-1m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(true, result.Item1.Get("f").AsBool);
            Assert.AreEqual(true, result.Item1.Get("g").AsBool);
            Assert.AreEqual(null, result.Item1.Get("h"));
        }

        [TestMethod]
        public void TestAndOperator()
        {
            string code = @"
#Const a = False And False
#Const b = False And True
#Const c = True And False
#Const d = True And True
#Const e = True And 5
#Const f = 1 And Null
#Const g = Null And 1
#Const h = Null And Null
#Const i = 0 And Null
#Const j = Null And 0
";
            var result = Preprocess(code);
            Assert.AreEqual(false, result.Item1.Get("a").AsBool);
            Assert.AreEqual(false, result.Item1.Get("b").AsBool);
            Assert.AreEqual(false, result.Item1.Get("c").AsBool);
            Assert.AreEqual(true, result.Item1.Get("d").AsBool);
            Assert.AreEqual(5m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("f"));
            Assert.AreEqual(null, result.Item1.Get("g"));
            Assert.AreEqual(null, result.Item1.Get("h"));
            Assert.AreEqual(0, result.Item1.Get("i").AsDecimal);
            Assert.AreEqual(0, result.Item1.Get("j").AsDecimal);
        }

        [TestMethod]
        public void TestGeqOperator()
        {
            string code = @"
#Const a = 2 >= 1
#Const b = 1 >= 1
#Const c = 0 >= 1
#Const d = False >= True
#Const e = False >= False
#Const f = True >= False
#Const g = ""b"" >= ""a""
#Const h = ""b"" >= ""b""
#Const i = ""a"" >= ""b""
#Const j = ""2"" >= 1
#Const k = ""2"" >= 2
#Const l = ""1"" >= 2
#Const m = #01-01-2000# >= #01-01-1900#
#Const n = #01-01-2000# >= #01-01-2000#
#Const o = #01-01-1900# >= #01-01-2000#
#Const p = Null >= Null
#Const q = 1 >= Null
#Const r = Null >= 1
#Const s = ""a"" >= Empty
#Const t = Empty >= Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(true, result.Item1.Get("a").AsBool);
            Assert.AreEqual(true, result.Item1.Get("b").AsBool);
            Assert.AreEqual(false, result.Item1.Get("c").AsBool);
            Assert.AreEqual(true, result.Item1.Get("d").AsBool);
            Assert.AreEqual(true, result.Item1.Get("e").AsBool);
            Assert.AreEqual(false, result.Item1.Get("f").AsBool);
            Assert.AreEqual(true, result.Item1.Get("g").AsBool);
            Assert.AreEqual(true, result.Item1.Get("h").AsBool);
            Assert.AreEqual(false, result.Item1.Get("i").AsBool);
            Assert.AreEqual(true, result.Item1.Get("j").AsBool);
            Assert.AreEqual(true, result.Item1.Get("k").AsBool);
            Assert.AreEqual(false, result.Item1.Get("l").AsBool);
            Assert.AreEqual(true, result.Item1.Get("m").AsBool);
            Assert.AreEqual(true, result.Item1.Get("n").AsBool);
            Assert.AreEqual(false, result.Item1.Get("o").AsBool);
            Assert.AreEqual(null, result.Item1.Get("p"));
            Assert.AreEqual(null, result.Item1.Get("q"));
            Assert.AreEqual(null, result.Item1.Get("r"));
            Assert.AreEqual(true, result.Item1.Get("s").AsBool);
            Assert.AreEqual(true, result.Item1.Get("t").AsBool);
        }

        [TestMethod]
        public void TestGtOperator()
        {
            string code = @"
#Const a = 2 > 1
#Const b = 1 > 1
#Const c = 0 > 1
#Const d = False > True
#Const e = False > False
#Const f = True > False
#Const g = ""b"" > ""a""
#Const h = ""b"" > ""b""
#Const i = ""a"" > ""b""
#Const j = ""2"" > 1
#Const k = ""2"" > 2
#Const l = ""1"" > 2
#Const m = #01-01-2000# > #01-01-1900#
#Const n = #01-01-2000# > #01-01-2000#
#Const o = #01-01-1900# > #01-01-2000#
#Const p = Null > Null
#Const q = 1 > Null
#Const r = Null > 1
#Const s = ""a"" > Empty
#Const t = Empty > Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(true, result.Item1.Get("a").AsBool);
            Assert.AreEqual(false, result.Item1.Get("b").AsBool);
            Assert.AreEqual(false, result.Item1.Get("c").AsBool);
            Assert.AreEqual(true, result.Item1.Get("d").AsBool);
            Assert.AreEqual(false, result.Item1.Get("e").AsBool);
            Assert.AreEqual(false, result.Item1.Get("f").AsBool);
            Assert.AreEqual(true, result.Item1.Get("g").AsBool);
            Assert.AreEqual(false, result.Item1.Get("h").AsBool);
            Assert.AreEqual(false, result.Item1.Get("i").AsBool);
            Assert.AreEqual(true, result.Item1.Get("j").AsBool);
            Assert.AreEqual(false, result.Item1.Get("k").AsBool);
            Assert.AreEqual(false, result.Item1.Get("l").AsBool);
            Assert.AreEqual(true, result.Item1.Get("m").AsBool);
            Assert.AreEqual(false, result.Item1.Get("n").AsBool);
            Assert.AreEqual(false, result.Item1.Get("o").AsBool);
            Assert.AreEqual(null, result.Item1.Get("p"));
            Assert.AreEqual(null, result.Item1.Get("q"));
            Assert.AreEqual(null, result.Item1.Get("r"));
            Assert.AreEqual(true, result.Item1.Get("s").AsBool);
            Assert.AreEqual(false, result.Item1.Get("t").AsBool);
        }

        [TestMethod]
        public void TestLeqOperator()
        {
            string code = @"
#Const a = 2 <= 1
#Const b = 1 <= 1
#Const c = 0 <= 1
#Const d = False <= True
#Const e = False <= False
#Const f = True <= False
#Const g = ""b"" <= ""a""
#Const h = ""b"" <= ""b""
#Const i = ""a"" <= ""b""
#Const j = ""2"" <= 1
#Const k = ""2"" <= 2
#Const l = ""1"" <= 2
#Const m = #01-01-2000# <= #01-01-1900#
#Const n = #01-01-2000# <= #01-01-2000#
#Const o = #01-01-1900# <= #01-01-2000#
#Const p = Null <= Null
#Const q = 1 <= Null
#Const r = Null <= 1
#Const s = ""a"" <= Empty
#Const t = Empty <= Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(false, result.Item1.Get("a").AsBool);
            Assert.AreEqual(true, result.Item1.Get("b").AsBool);
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
            Assert.AreEqual(false, result.Item1.Get("d").AsBool);
            Assert.AreEqual(true, result.Item1.Get("e").AsBool);
            Assert.AreEqual(true, result.Item1.Get("f").AsBool);
            Assert.AreEqual(false, result.Item1.Get("g").AsBool);
            Assert.AreEqual(true, result.Item1.Get("h").AsBool);
            Assert.AreEqual(true, result.Item1.Get("i").AsBool);
            Assert.AreEqual(false, result.Item1.Get("j").AsBool);
            Assert.AreEqual(true, result.Item1.Get("k").AsBool);
            Assert.AreEqual(true, result.Item1.Get("l").AsBool);
            Assert.AreEqual(false, result.Item1.Get("m").AsBool);
            Assert.AreEqual(true, result.Item1.Get("n").AsBool);
            Assert.AreEqual(true, result.Item1.Get("o").AsBool);
            Assert.AreEqual(null, result.Item1.Get("p"));
            Assert.AreEqual(null, result.Item1.Get("q"));
            Assert.AreEqual(null, result.Item1.Get("r"));
            Assert.AreEqual(false, result.Item1.Get("s").AsBool);
            Assert.AreEqual(true, result.Item1.Get("t").AsBool);
        }

        [TestMethod]
        public void TestLtOperator()
        {
            string code = @"
#Const a = 2 < 1
#Const b = 1 < 1
#Const c = 0 < 1
#Const d = False < True
#Const e = False < False
#Const f = True < False
#Const g = ""b"" < ""a""
#Const h = ""b"" < ""b""
#Const i = ""a"" < ""b""
#Const j = ""2"" < 1
#Const k = ""2"" < 2
#Const l = ""1"" < 2
#Const m = #01-01-2000# < #01-01-1900#
#Const n = #01-01-2000# < #01-01-2000#
#Const o = #01-01-1900# < #01-01-2000#
#Const p = Null < Null
#Const q = 1 < Null
#Const r = Null < 1
#Const s = ""a"" < Empty
#Const t = Empty < Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(false, result.Item1.Get("a").AsBool);
            Assert.AreEqual(false, result.Item1.Get("b").AsBool);
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
            Assert.AreEqual(false, result.Item1.Get("d").AsBool);
            Assert.AreEqual(false, result.Item1.Get("e").AsBool);
            Assert.AreEqual(true, result.Item1.Get("f").AsBool);
            Assert.AreEqual(false, result.Item1.Get("g").AsBool);
            Assert.AreEqual(false, result.Item1.Get("h").AsBool);
            Assert.AreEqual(true, result.Item1.Get("i").AsBool);
            Assert.AreEqual(false, result.Item1.Get("j").AsBool);
            Assert.AreEqual(false, result.Item1.Get("k").AsBool);
            Assert.AreEqual(true, result.Item1.Get("l").AsBool);
            Assert.AreEqual(false, result.Item1.Get("m").AsBool);
            Assert.AreEqual(false, result.Item1.Get("n").AsBool);
            Assert.AreEqual(true, result.Item1.Get("o").AsBool);
            Assert.AreEqual(null, result.Item1.Get("p"));
            Assert.AreEqual(null, result.Item1.Get("q"));
            Assert.AreEqual(null, result.Item1.Get("r"));
            Assert.AreEqual(false, result.Item1.Get("s").AsBool);
            Assert.AreEqual(false, result.Item1.Get("t").AsBool);
        }

        [TestMethod]
        public void TestEqOperator()
        {
            string code = @"
#Const a = 2 = 1
#Const b = 1 = 1
#Const c = False = True
#Const d = False = False
#Const e = ""b"" = ""a""
#Const f = ""b"" = ""b""
#Const g = ""2"" = 1
#Const h = ""2"" = 2
#Const i = #01-01-2000# = #01-01-1900#
#Const j = #01-01-2000# = #01-01-2000#
#Const k = Null = Null
#Const l = 1 = Null
#Const m = Null = 1
#Const n = ""a"" = Empty
#Const o = Empty = Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(false, result.Item1.Get("a").AsBool);
            Assert.AreEqual(true, result.Item1.Get("b").AsBool);
            Assert.AreEqual(false, result.Item1.Get("c").AsBool);
            Assert.AreEqual(true, result.Item1.Get("d").AsBool);
            Assert.AreEqual(false, result.Item1.Get("e").AsBool);
            Assert.AreEqual(true, result.Item1.Get("f").AsBool);
            Assert.AreEqual(false, result.Item1.Get("g").AsBool);
            Assert.AreEqual(true, result.Item1.Get("h").AsBool);
            Assert.AreEqual(false, result.Item1.Get("i").AsBool);
            Assert.AreEqual(true, result.Item1.Get("j").AsBool);
            Assert.AreEqual(null, result.Item1.Get("k"));
            Assert.AreEqual(null, result.Item1.Get("l"));
            Assert.AreEqual(null, result.Item1.Get("m"));
            Assert.AreEqual(false, result.Item1.Get("n").AsBool);
            Assert.AreEqual(true, result.Item1.Get("o").AsBool);
        }

        [TestMethod]
        public void TestNeqOperator()
        {
            string code = @"
#Const a = 2 <> 1
#Const b = 1 <> 1
#Const c = False <> True
#Const d = False <> False
#Const e = ""b"" <> ""a""
#Const f = ""b"" <> ""b""
#Const g = ""2"" <> 1
#Const h = ""2"" <> 2
#Const i = #01-01-2000# <> #01-01-1900#
#Const j = #01-01-2000# <> #01-01-2000#
#Const k = Null <> Null
#Const l = 1 <> Null
#Const m = Null <> 1
#Const n = ""a"" <> Empty
#Const o = Empty <> Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(true, result.Item1.Get("a").AsBool);
            Assert.AreEqual(false, result.Item1.Get("b").AsBool);
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
            Assert.AreEqual(false, result.Item1.Get("d").AsBool);
            Assert.AreEqual(true, result.Item1.Get("e").AsBool);
            Assert.AreEqual(false, result.Item1.Get("f").AsBool);
            Assert.AreEqual(true, result.Item1.Get("g").AsBool);
            Assert.AreEqual(false, result.Item1.Get("h").AsBool);
            Assert.AreEqual(true, result.Item1.Get("i").AsBool);
            Assert.AreEqual(false, result.Item1.Get("j").AsBool);
            Assert.AreEqual(null, result.Item1.Get("k"));
            Assert.AreEqual(null, result.Item1.Get("l"));
            Assert.AreEqual(null, result.Item1.Get("m"));
            Assert.AreEqual(true, result.Item1.Get("n").AsBool);
            Assert.AreEqual(false, result.Item1.Get("o").AsBool);
        }

        [TestMethod]
        public void TestConcatOperator()
        {
            string code = @"
#Const a = True & ""a""
#Const b = Null & Null
#Const c = 1 & Null
#Const d = Null & 1
#Const e = #01-01-1900# & 1
#Const f = 1 & Empty
#Const g = Empty & 1
#Const h = Empty & Empty
";
            var result = Preprocess(code);
            Assert.AreEqual("Truea", result.Item1.Get("a").AsString);
            Assert.AreEqual(null, result.Item1.Get("b"));
            Assert.AreEqual("1", result.Item1.Get("c").AsString);
            Assert.AreEqual("1", result.Item1.Get("d").AsString);
            Assert.AreEqual(new DateTime(1900, 1, 1).ToShortDateString() + "1", result.Item1.Get("e").AsString);
            Assert.AreEqual("1", result.Item1.Get("f").AsString);
            Assert.AreEqual("1", result.Item1.Get("g").AsString);
            Assert.AreEqual(string.Empty, result.Item1.Get("h").AsString);
        }

        [TestMethod]
        public void TestPowOperator()
        {
            string code = @"
#Const a = 2 ^ 3
#Const b = Empty ^ False
#Const c = 0 ^ #30-12-1899#
#Const d = Null ^ 3
#Const e = 2 ^ Null
#Const f = Null ^ Null
";
            var result = Preprocess(code);
            Assert.AreEqual(8m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(1m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(1m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("d"));
            Assert.AreEqual(null, result.Item1.Get("e"));
            Assert.AreEqual(null, result.Item1.Get("f"));
        }

        [TestMethod]
        public void TestModOperator()
        {
            string code = @"
#Const a = 2 Mod True
#Const b = 10 Mod 3
#Const c = 3 Mod #1/1/1900#
#Const d = Null Mod 3
#Const e = 2 Mod Null
#Const f = Null Mod Null
";
            var result = Preprocess(code);
            Assert.AreEqual(0m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(1m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(1m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("d"));
            Assert.AreEqual(null, result.Item1.Get("e"));
            Assert.AreEqual(null, result.Item1.Get("f"));
        }

        [TestMethod]
        public void TestIntDivOperator()
        {
            string code = @"
#Const a = 5 \ 2
#Const b = 5.1 \ 2
#Const c = 4.9 \ 2
#Const d = -5 \ 2
#Const e = -5.1 \ 2
#Const f = -4.9 \ 2
#Const g = Null \ 3
#Const h = 2 \ Null
#Const i = Null \ Null
";
            var result = Preprocess(code);
            Assert.AreEqual(2m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(2m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(2m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(-2m, result.Item1.Get("d").AsDecimal);
            Assert.AreEqual(-2m, result.Item1.Get("e").AsDecimal);
            Assert.AreEqual(-2m, result.Item1.Get("f").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("g"));
            Assert.AreEqual(null, result.Item1.Get("h"));
            Assert.AreEqual(null, result.Item1.Get("i"));
        }

        [TestMethod]
        public void TestMultOperator()
        {
            string code = @"
#Const a = 5.5 * 2
#Const b = 5.5 * True
#Const c = Null * 3
#Const d = 2 * Null
#Const e = Null * Null
";
            var result = Preprocess(code);
            Assert.AreEqual(11m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(-5.5m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("c"));
            Assert.AreEqual(null, result.Item1.Get("d"));
            Assert.AreEqual(null, result.Item1.Get("e"));
        }

        [TestMethod]
        public void TestDivOperator()
        {
            string code = @"
#Const a = 5.5 / 2
#Const b = 5.5 / True
#Const c = Null / 3
#Const d = 2 / Null
#Const e = Null / Null
";
            var result = Preprocess(code);
            Assert.AreEqual(2.75m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(-5.5m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(null, result.Item1.Get("c"));
            Assert.AreEqual(null, result.Item1.Get("d"));
            Assert.AreEqual(null, result.Item1.Get("e"));
        }

        [TestMethod]
        public void TestStringLiteral()
        {
            string code = @"
#Const a = ""abc""
#Const b = ""a""""b""
";
            var result = Preprocess(code);
            Assert.AreEqual("abc", result.Item1.Get("a").AsString);
            Assert.AreEqual("a\"\"b", result.Item1.Get("b").AsString);
        }

        [TestMethod]
        public void TestNumberLiteral()
        {
            string code = @"
#Const a = &HAF%
#Const b = &O423^
#Const c = -50.323e5
";
            var result = Preprocess(code);
            Assert.AreEqual(175m, result.Item1.Get("a").AsDecimal);
            Assert.AreEqual(275m, result.Item1.Get("b").AsDecimal);
            Assert.AreEqual(-5032300m, result.Item1.Get("c").AsDecimal);
        }

        [TestMethod]
        public void TestDateLiteral()
        {
            string code = @"
#Const a = #30-12-1900#
#Const b = #12-30-1900#
#Const c = #12-30#
#Const d = #12:00#
#Const e = #mar-1999#
#Const f = #2pm#
#Const g = #12-1999#
#Const h = #1999-11#
#Const i = #2010-may#
#Const j = #2010-july-15#
#Const k = #15:14:13#
#Const l = #15:14:13am#
#Const m = #12am#
#Const n = #12:13#
";
            var result = Preprocess(code);
            Assert.AreEqual(new DateTime(1900, 12, 30), result.Item1.Get("a").AsDate);
            Assert.AreEqual(new DateTime(1900, 12, 30), result.Item1.Get("b").AsDate);
            Assert.AreEqual(new DateTime(DateTime.Now.Year, 12, 30), result.Item1.Get("c").AsDate);
            Assert.AreEqual(new DateTime(1899, 12, 30, 12, 0, 0), result.Item1.Get("d").AsDate);
            Assert.AreEqual(new DateTime(1999, 3, 1), result.Item1.Get("e").AsDate);
            Assert.AreEqual(new DateTime(1899, 12, 30, 14, 0, 0), result.Item1.Get("f").AsDate);
            Assert.AreEqual(new DateTime(1999, 12, 1, 0, 0, 0), result.Item1.Get("g").AsDate);
            Assert.AreEqual(new DateTime(1999, 11, 1, 0, 0, 0), result.Item1.Get("h").AsDate);
            Assert.AreEqual(new DateTime(2010, 5, 1, 0, 0, 0), result.Item1.Get("i").AsDate);
            Assert.AreEqual(new DateTime(2010, 7, 15, 0, 0, 0), result.Item1.Get("j").AsDate);
            Assert.AreEqual(new DateTime(1899, 12, 30, 15, 14, 13), result.Item1.Get("k").AsDate);
            // "A <ampm> element has no significance if the <hour-value> is greater than 12."
            Assert.AreEqual(new DateTime(1899, 12, 30, 15, 14, 13), result.Item1.Get("l").AsDate);
            Assert.AreEqual(new DateTime(1899, 12, 30, 0, 0, 0), result.Item1.Get("m").AsDate);
            Assert.AreEqual(new DateTime(1899, 12, 30, 12, 13, 0), result.Item1.Get("n").AsDate);
        }

        [TestMethod]
        public void TestKeywordLiterals()
        {
            string code = @"
#Const a = true
#Const b = false
#Const c = Nothing
#Const d = Null
#Const e = Empty
";
            var result = Preprocess(code);
            Assert.AreEqual(true, result.Item1.Get("a").AsBool);
            Assert.AreEqual(false, result.Item1.Get("b").AsBool);
            Assert.AreEqual(null, result.Item1.Get("c"));
            Assert.AreEqual(null, result.Item1.Get("d"));
            Assert.AreEqual(EmptyValue.Value, result.Item1.Get("e"));
        }

        [TestMethod]
        public void TestComplexExpressions()
        {
            string code = @"
#Const a = 23
#Const b = 500
#Const c = a >= b Or True ^ #1/1/1900#
#Const d = True + #1/1/1800# - (4 * Empty Mod (Abs(-5)))
";
            var result = Preprocess(code);
            Assert.AreEqual(1m, result.Item1.Get("c").AsDecimal);
            Assert.AreEqual(new DateTime(1799, 12, 31), result.Item1.Get("d").AsDate);
        }

        [TestMethod]
        public void TestOperatorPrecedence()
        {
            string code = @"
#Const a = 2 ^ 3 + -5 * 40 / 2 \ 4 Mod 2 + 3 - (2 * 4) Xor 4 Eqv 5 Imp 6 Or 2 And True
";
            var result = Preprocess(code);
            Assert.AreEqual(7m, result.Item1.Get("a").AsDecimal);
        }

        [TestMethod]
        public void TestLocaleJapanese()
        {
            string code = @"
#Const a = CDate(""2016/03/02"")
#Const b = CBool(""tRuE"")
#Const c = CBool(""#TRUE#"")
#Const d = ""ß"" = ""ss""
";
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("ja-jp");
            var result = Preprocess(code);
            Assert.AreEqual(new DateTime(2016, 3, 2), result.Item1.Get("a").AsDate);
            Assert.AreEqual(true, result.Item1.Get("b").AsBool);
            Assert.AreEqual(true, result.Item1.Get("c").AsBool);
            Assert.AreEqual(true, result.Item1.Get("d").AsBool);
        }

        [TestMethod]
        public void TestLocaleGerman()
        {
            string code = @"
#Const a = 82.5235
";
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("de-de");
            var result = Preprocess(code);
            Assert.AreEqual(82.5235m, result.Item1.Get("a").AsDecimal);
        }

        [TestMethod]
        public void TestPreprocessingLiveDeadCode()
        {
            string code = @"
#Const a = 2 + 5

#If a = 7 Then
    Public Sub Alive()
        #If True Then
            Debug.Print 2
        #ElseIf 1 = 2 Then
            Debug.Print 4
        #End If
    End Sub
#Else
    Public Sub Dead()
        Debug.Print 3
    End Sub
#End If
";

            string evaluated = @"



    Public Sub Alive()

            Debug.Print 2



    End Sub





";
            var result = Preprocess(code);
            Assert.AreEqual(evaluated, result.Item2.AsString);
        }

        [TestMethod]
        public void TestPreprocessingLiveDeadCodeTokensDoNotGetRemoved()
        {
            string code = @"
#Const a = 2 + 5

#If a = 7 Then
    Public Sub Alive()
        #If True Then
            Debug.Print 2
        #ElseIf 1 = 2 Then
            Debug.Print 4
        #End If
    End Sub
#Else
    Public Sub Dead()
        Debug.Print 3
    End Sub
#End If
";

            string evaluated = @"
#Const a = 2 + 5

#If a = 7 Then
    Public Sub Alive()
        #If True Then
            Debug.Print 2
        #ElseIf 1 = 2 Then
            Debug.Print 4
        #End If
    End Sub
#Else
    Public Sub Dead()
        Debug.Print 3
    End Sub
#End If
";
            var result = Preprocess(code);
            var allTokenText = TokenText(result.Item2.AsTokens);
            Assert.AreEqual(evaluated, allTokenText);
        }

        [TestMethod]
        public void TestPreprocessingNoConditionalCompilation()
        {
            string code = @"
Public Sub Unchanged()
    Debug.Print 3
a:
    Debug.Print 5
End Sub
";

            string evaluated = @"
Public Sub Unchanged()
    Debug.Print 3
a:
    Debug.Print 5
End Sub
";
            var result = Preprocess(code);
            Assert.AreEqual(evaluated, result.Item2.AsString);
        }

        [TestMethod]
        public void TestLogicalLinesHasConditionalCompilationKeywords()
        {
            string code = @"
Sub FileTest()
    Open ""TESTFILE"" For Input As #iFile
    Close #iFile
End Sub
";

            string evaluated = @"
Sub FileTest()
    Open ""TESTFILE"" For Input As #iFile
    Close #iFile
End Sub
";
            var result = Preprocess(code);
            Assert.AreEqual(evaluated, result.Item2.AsString);
        }

        [TestMethod]
        public void TestPtrSafeKeywordAsConstant()
        {
            string code = @"
#Const PtrSafe = True
#If PtrSafe Then
    Public Declare PtrSafe Function GetActiveWindow Lib ""User32"" () As LongPtr
#Else
    Public Declare Function GetActiveWindow Lib ""User32""() As Long
#End If
";

            string evaluated = @"


    Public Declare PtrSafe Function GetActiveWindow Lib ""User32"" () As LongPtr



";
            var result = Preprocess(code);
            Assert.AreEqual(evaluated, result.Item2.AsString);
        }

        [TestMethod]
        public void TestIgnoresComment()
        {
            string code = @"
       ' #if defined(WIN32)
        '    unsigned long cbElements;   // Size of an element of the array.
        '                                // Does not include size of
        '                                // pointed-to data.
        '    unsigned long cLocks;       // Number of times the array has been
        '                                // locked without corresponding unlock.
        ' #Else
        '    unsigned short cbElements;
        '    unsigned short cLocks;
        '    unsigned long handle;       // Used on Macintosh only.
        ' #End If
";

            string evaluated = @"
       ' #if defined(WIN32)
        '    unsigned long cbElements;   // Size of an element of the array.
        '                                // Does not include size of
        '                                // pointed-to data.
        '    unsigned long cLocks;       // Number of times the array has been
        '                                // locked without corresponding unlock.
        ' #Else
        '    unsigned short cbElements;
        '    unsigned short cLocks;
        '    unsigned long handle;       // Used on Macintosh only.
        ' #End If
";
            var result = Preprocess(code);
            Assert.AreEqual(evaluated, result.Item2.AsString);
        }

        private Tuple<SymbolTable<string, IValue>, IValue> Preprocess(string code)
        {
            SymbolTable<string, IValue> symbolTable = new SymbolTable<string, IValue>();
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAConditionalCompilationParser(tokens);
            parser.ErrorHandler = new BailErrorStrategy();
            //parser.AddErrorListener(new ExceptionErrorListener());
            var tree = parser.compilationUnit();
            var evaluator = new VBAPreprocessorVisitor(symbolTable, new VBAPredefinedCompilationConstants(7.01), tree.start.InputStream, tokens);
            var expr = evaluator.Visit(tree);
            var resultValue = expr.Evaluate();

            Debug.Assert(parser.NumberOfSyntaxErrors == 0);
            return Tuple.Create(symbolTable, resultValue);
        }

        private string TokenText(IEnumerable<IToken> tokens)
        {
            var builder = new StringBuilder();
            foreach(var token in tokens)
            {
                builder.Append(token.Text);
            }
            var withoutEOF = builder.ToString();
            while (withoutEOF.Length >= 5 && String.Equals(withoutEOF.Substring(withoutEOF.Length - 5, 5), "<EOF>"))
            {
                withoutEOF = withoutEOF.Substring(0, withoutEOF.Length - 5);
            }
            return withoutEOF;
        }
    }
}
