using System;
using System.Globalization;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class DateLiteralExpression : Expression
    {
        private readonly IExpression _tokenText;

        public DateLiteralExpression(IExpression tokenText)
        {
            _tokenText = tokenText;
        }

        public override IValue Evaluate()
        {
            string literal = _tokenText.Evaluate().AsString;
            var parser = new VBADateLiteralParser();
            var dateLiteral = parser.Parse(literal);
            var dateOrTime = dateLiteral.dateOrTime();
            int year;
            int month;
            int day;
            int hours;
            int mins;
            int seconds;

            Predicate<int> legalMonth = (x) => x >= 0 && x <= 12;
            Func<int, int, int, bool> legalDay = (m, d, y) =>
            {
                bool legalYear = y >= 0 && y <= 32767;
                bool legalM = legalMonth(m);
                bool legalD = false;
                if (legalYear && legalM)
                {
                    int daysInMonth = DateTime.DaysInMonth(y, m);
                    legalD = d >= 1 && d <= daysInMonth;
                }
                return legalYear && legalM && legalD;
            };

            Func<int, int> yearFunc = (x) =>
            {
                if (x >= 0 && x <= 29)
                {
                    return x + 2000;
                }
                else if (x >= 30 && x <= 99)
                {
                    return x + 1900;
                }
                else
                {
                    return x;
                }
            };

            int CY = DateTime.Now.Year;

            if (dateOrTime.dateValue() == null)
            {
                year = 1899;
                month = 12;
                day = 30;
            }
            else
            {
                var dateValue = dateOrTime.dateValue();
                var txt = dateOrTime.GetText();
                var L = dateValue.dateValuePart()[0];
                var M = dateValue.dateValuePart()[1];
                VBADateParser.DateValuePartContext R = null;
                if (dateValue.dateValuePart().Count == 3)
                {
                    R = dateValue.dateValuePart()[2];
                }
                // "If L and M are numbers and R is not present:"
                if (L.dateValueNumber() != null && M.dateValueNumber() != null && R == null)
                {
                    var LNumber = int.Parse(L.GetText(), CultureInfo.InvariantCulture);
                    var MNumber = int.Parse(M.GetText(), CultureInfo.InvariantCulture);
                    if (legalMonth(LNumber) && legalDay(LNumber, MNumber, CY))
                    {
                        month = LNumber;
                        day = MNumber;
                        year = CY;
                    }
                    else if ((legalMonth(MNumber) && legalDay(MNumber, LNumber, CY)))
                    {
                        month = MNumber;
                        day = LNumber;
                        year = CY;
                    }
                    else if (legalMonth(LNumber))
                    {
                        month = LNumber;
                        day = 1;
                        year = MNumber;
                    }
                    else if (legalMonth(MNumber))
                    {
                        month = MNumber;
                        day = 1;
                        year = LNumber;
                    }
                    else
                    {
                        throw new Exception("Invalid date: " + dateLiteral.GetText());
                    }
                }
                // "If L, M, and R are numbers:"
                else if (L.dateValueNumber() != null && M.dateValueNumber() != null && R != null && R.dateValueNumber() != null)
                {
                    var LNumber = int.Parse(L.GetText(), CultureInfo.InvariantCulture);
                    var MNumber = int.Parse(M.GetText(), CultureInfo.InvariantCulture);
                    var RNumber = int.Parse(R.GetText(), CultureInfo.InvariantCulture);
                    if (legalMonth(LNumber) && legalDay(LNumber, MNumber, yearFunc(RNumber)))
                    {
                        month = LNumber;
                        day = MNumber;
                        year = yearFunc(RNumber);
                    }
                    else if (legalMonth(MNumber) && legalDay(MNumber, RNumber, yearFunc(LNumber)))
                    {
                        month = MNumber;
                        day = RNumber;
                        year = yearFunc(LNumber);
                    }
                    else if (legalMonth(MNumber) && legalDay(MNumber, LNumber, yearFunc(RNumber)))
                    {
                        month = MNumber;
                        day = LNumber;
                        year = yearFunc(RNumber);
                    }
                    else
                    {
                        throw new Exception("Invalid date: " + dateLiteral.GetText());
                    }
                }
                // "If either L or M is not a number and R is not present:"
                else if ((L.dateValueNumber() == null || M.dateValueNumber() == null) && R == null)
                {
                    int N;
                    string monthName;
                    if (L.dateValueNumber() != null)
                    {
                        N = int.Parse(L.GetText(), CultureInfo.InvariantCulture);
                        monthName = M.GetText();
                    }
                    else
                    {
                        N = int.Parse(M.GetText(), CultureInfo.InvariantCulture);
                        monthName = L.GetText();
                    }
                    int monthNameNumber;
                    if (monthName.Length == 3)
                    {
                        monthNameNumber = DateTime.ParseExact(monthName, "MMM", CultureInfo.InvariantCulture).Month;
                    }
                    else
                    {
                        monthNameNumber = DateTime.ParseExact(monthName, "MMMM", CultureInfo.InvariantCulture).Month;
                    }
                    if (legalDay(monthNameNumber, N, CY))
                    {
                        month = monthNameNumber;
                        day = N;
                        year = CY;
                    }
                    else
                    {
                        month = monthNameNumber;
                        day = 1;
                        year = N;
                    }
                }
                // "Otherwise, R is present and one of L, M, and R is not a number:"
                else
                {
                    int N1;
                    int N2;
                    string monthName;
                    if (L.dateValueNumber() == null)
                    {
                        monthName = L.GetText();
                        N1 = int.Parse(M.GetText(), CultureInfo.InvariantCulture);
                        N2 = int.Parse(R.GetText(), CultureInfo.InvariantCulture);
                    }
                    else if (M.dateValueNumber() == null)
                    {
                        monthName = M.GetText();
                        N1 = int.Parse(L.GetText(), CultureInfo.InvariantCulture);
                        N2 = int.Parse(R.GetText(), CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        monthName = R.GetText();
                        N1 = int.Parse(L.GetText(), CultureInfo.InvariantCulture);
                        N2 = int.Parse(M.GetText(), CultureInfo.InvariantCulture);
                    }
                    int monthNameNumber;
                    if (monthName.Length == 3)
                    {
                        monthNameNumber = DateTime.ParseExact(monthName, "MMM", CultureInfo.InvariantCulture).Month;
                    }
                    else
                    {
                        monthNameNumber = DateTime.ParseExact(monthName, "MMMM", CultureInfo.InvariantCulture).Month;
                    }
                    if (legalDay(monthNameNumber, N1, yearFunc(N2)))
                    {
                        month = monthNameNumber;
                        day = N1;
                        year = yearFunc(N2);
                    }
                    else if (legalDay(monthNameNumber, N2, yearFunc(N1)))
                    {
                        month = monthNameNumber;
                        day = N2;
                        year = yearFunc(N1);
                    }
                    else
                    {
                        throw new Exception("Invalid date: " + dateLiteral.GetText());
                    }
                }
            }

            if (dateOrTime.timeValue() == null)
            {
                hours = 0;
                mins = 0;
                seconds = 0;
            }
            else
            {
                var timeValue = dateOrTime.timeValue();
                hours = int.Parse(timeValue.timeValuePart()[0].GetText(), CultureInfo.InvariantCulture);
                if (timeValue.timeValuePart().Count == 1)
                {
                    mins = 0;
                }
                else
                {
                    mins = int.Parse(timeValue.timeValuePart()[1].GetText(), CultureInfo.InvariantCulture);
                }
                if (timeValue.timeValuePart().Count < 3)
                {
                    seconds = 0;
                }
                else
                {
                    seconds = int.Parse(timeValue.timeValuePart()[2].GetText(), CultureInfo.InvariantCulture);
                }
                var amPm = timeValue.AMPM();
                if (amPm != null && (amPm.GetText().ToUpper() == "P" || amPm.GetText().ToUpper() == "PM") && hours >= 0 && hours <= 11)
                {
                    hours += 12;
                }
                else if (amPm != null && (amPm.GetText().ToUpper() == "A" || amPm.GetText().ToUpper() == "AM") && hours == 12)
                {
                    hours = 0;
                }
            }
            return new DateValue(new DateTime(year, month, day, hours, mins, seconds));
        }
    }
}
