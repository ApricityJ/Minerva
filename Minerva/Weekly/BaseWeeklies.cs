using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Weekly
{
    using Department;

    class BaseWeeklies : AbstractWeeklies
    {

        public List<WeeklyItem> CurrentWeekWorks { get; set; }
        public List<WeeklyItem> UnnormalCases { get; set; }

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
                .ForEach(w => CurrentWeekWorks.Concat(w.Value.CurrentWeekWork));

            WeeklySet.ToList()
                .ForEach(w => UnnormalCases.Concat(w.Value.UnnormalCase));

            return this;
        }


        public override AbstractWeeklies Sort()
        {
            CurrentWeekWorks.Sort();
            UnnormalCases.Sort();

            return this;
        }



        public override AbstractWeeklies Save()
        {
            throw new NotImplementedException();
        }
    }
}
