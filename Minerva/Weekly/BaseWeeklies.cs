using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Weekly
{
    using Department;

    /// <summary>
    /// 周报集合的基础类，用于处理周报的集合
    /// </summary>
    class BaseWeeklies : AbstractWeeklies
    {

        public List<WeeklyItem> CurrentWeekWorks { get; set; }
        public List<WeeklyItem> UnnormalCases { get; set; }

        public BaseWeeklies()
        {
            WeeklySet = new Dictionary<InnerDepartment, BaseWeekly>();

            CurrentWeekWorks = new List<WeeklyItem>();
            UnnormalCases = new List<WeeklyItem>();
        }

        public override AbstractWeeklies Load()
        {
            foreach (InnerDepartment department in Enum.GetValues(typeof(InnerDepartment)))
            {
                if (!department.ToString().StartsWith("Dev"))
                {
                    WeeklySet.Add(department, new DevWeekly(Env.Instance.ToWeeklyPath(department)));
                }
            }
            return this;
        }


        public override AbstractWeeklies Summarize()
        {
            WeeklySet.ToList()
                .ForEach(w => CurrentWeekWorks = CurrentWeekWorks.Concat(w.Value.CurrentWeekWork).ToList());

            WeeklySet.ToList()
                .ForEach(w => UnnormalCases = UnnormalCases.Concat(w.Value.UnnormalCase).ToList());

            return this;
        }


        public override AbstractWeeklies Sort()
        {
            CurrentWeekWorks.Sort();
            UnnormalCases.Sort();

            return this;
        }

    }
}
