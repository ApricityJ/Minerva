using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Util
{
    class DateUtil
    {
        public static string ToCurrentYear()
        {
            return DateTime.Now.Year.ToString();
        }

        public static int ToWeekOfYear()
        {
            CultureInfo ci = new CultureInfo("zh-CN");
            System.Globalization.Calendar cal = ci.Calendar;
            CalendarWeekRule rule = ci.DateTimeFormat.CalendarWeekRule;
            DayOfWeek dow = DayOfWeek.Monday;
            return cal.GetWeekOfYear(DateTime.Now, rule, dow);
        }

        public static string ToWeekEnd()
        {
            return DateTime.Now
                .AddDays(DayOfWeek.Friday - DateTime.Now.DayOfWeek)
                .ToString("M月dd日");
        }

        public static string ToWeekBegin()
        {
            return DateTime.Now
                .AddDays(DayOfWeek.Monday - DateTime.Now.DayOfWeek)
                .ToString("M月dd日");
        }
    }
}
