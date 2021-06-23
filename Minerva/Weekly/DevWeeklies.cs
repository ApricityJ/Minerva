using System;
using System.Collections.Generic;
using System.Linq;


namespace Minerva.Weekly
{

    using Department;

    /// <summary>
    /// 开发周报类
    /// 用于处理周报中项目类内容
    /// 在第一步，提取周报与项目计划比较中使用
    /// </summary>

    class DevWeeklies : BaseWeeklies
    {

        public List<WeeklyItem> CurrentWeekReleases { get; set; }
        public List<WeeklyItem> NextWeekReleasePlans { get; set; }
        public List<WeeklyItem> ProjectWorkList { get; set; }


        public DevWeeklies()
        {
            WeeklySet = new Dictionary<InnerDepartment, BaseWeekly>();

            CurrentWeekWorks = new List<WeeklyItem>();
            UnnormalCases = new List<WeeklyItem>();
            CurrentWeekReleases = new List<WeeklyItem>();
            NextWeekReleasePlans = new List<WeeklyItem>();

            ProjectWorkList = new List<WeeklyItem>();
        }


        //加载周报
        public override AbstractWeeklies Load()
        {
            foreach (InnerDepartment department in Enum.GetValues(typeof(InnerDepartment)))
            {
                if (department.ToString().StartsWith("Dev"))
                {
                    WeeklySet.Add(department, new DevWeekly(Env.Instance.ToWeeklyPath(department)));
                }
            }
            return this;
        }

        //筛选项目相关的工作内容
        public override AbstractWeeklies Summarize()
        {

            WeeklySet.ToList()
              .ForEach(w => CurrentWeekWorks = CurrentWeekWorks.Concat(w.Value.CurrentWeekWork).ToList());

            WeeklySet.ToList()
                .ForEach(w => UnnormalCases = UnnormalCases.Concat(w.Value.UnnormalCase).ToList());

            WeeklySet.ToList()
                .ForEach(w => CurrentWeekReleases = CurrentWeekReleases.Concat(((DevWeekly)w.Value).CurrentWeekRelease).ToList());

            WeeklySet.ToList()
                .ForEach(w => NextWeekReleasePlans = NextWeekReleasePlans.Concat(((DevWeekly)w.Value).NextWeekReleasePlan).ToList());


            //筛选项目类工作（用于与工作计划交互匹配）
            ProjectWorkList = CurrentWeekWorks.Where(w => w.IsProjectWork())
                .ToList();

            return this;
        }

        //判断一个项目名称是否出现在周报中
        //用于寻找 [项目计划] 中存在 但 [周报] 中不存在的项目
        public bool IsExist(string projectName)
        {
            return ProjectWorkList.Any(item => item.Name.Trim().Equals(projectName.Trim()));
        }

        public override AbstractWeeklies Sort()
        {
            CurrentWeekWorks.Sort();
            UnnormalCases.Sort();
            CurrentWeekReleases.Sort();
            NextWeekReleasePlans.Sort();

            return this;
        }

        public override AbstractWeeklies Save()
        {
            throw new NotImplementedException();
        }
    }
}
