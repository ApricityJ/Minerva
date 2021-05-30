using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Minerva.Weekly
{
    using Department;
    using Util;

    class BaseWeeklies : AbstractWeelies
    {

        public List<WeeklyItem> CurrentWeekWorks { get; set; }
        public List<WeeklyItem> UnnormalCases { get; set; }
        public List<WeeklyItem> CurrentWeekReleases { get; set; }
        public List<WeeklyItem> NextWeekReleasePlans { get; set; }

        private string template = "科技与产品管理部周报(#Year#)年第(#Week#)期.et";
        private string targetPath;

        public BaseWeeklies()
        {
            WeeklyCollection = new Dictionary<InnerDepartment, BaseWeekly>();
            CurrentWeekWorks = new List<WeeklyItem>();
            UnnormalCases = new List<WeeklyItem>();
            CurrentWeekReleases = new List<WeeklyItem>();
            NextWeekReleasePlans = new List<WeeklyItem>();
        }

        private void ToTargetPath()
        {
            targetPath = template.Replace("#Year#", DateUtil.ToCurrentYear())
                .Replace("#Week#", DateUtil.ToWeekOfYear().ToString());
            targetPath = Path.Combine(Env.Instance.RootDir, targetPath);

            if (File.Exists(targetPath))
            {
                File.Delete(targetPath);
            }
        }

        public override AbstractWeelies LoadWeekly()
        {
            foreach (InnerDepartment department in Enum.GetValues(typeof(InnerDepartment)))
            {
                WeeklyCollection.Add(department, new BaseWeekly(Env.Instance.ToWeeklyPath(department)));
            }
            return this;
        }

        public override AbstractWeelies Summarize()
        {

            WeeklyCollection.ToList()
                .ForEach(w => CurrentWeekWorks.Concat(w.Value.CurrentWeekWork));

            WeeklyCollection.ToList()
                .ForEach(w => UnnormalCases.Concat(w.Value.UnnormalCase));

            WeeklyCollection.ToList()
                .Where(w=>w.Key.ToString().StartsWith("Dev"))
                .ToList()
                .ForEach(w => CurrentWeekReleases.Concat(((DevWeekly)w.Value).CurrentWeekRelease));

            WeeklyCollection.ToList()
                .Where(w => w.Key.ToString().StartsWith("Dev"))
                .ToList()
                .ForEach(w => NextWeekReleasePlans.Concat(((DevWeekly)w.Value).NextWeekReleasePlan));


            CurrentWeekReleases.Sort();
            NextWeekReleasePlans.Sort();

            return this;
        }
    }
}
