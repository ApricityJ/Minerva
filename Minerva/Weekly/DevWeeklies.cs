using System;
using System.Collections.Generic;
using System.Linq;


namespace Minerva.Weekly
{

    using Department;

    /// <summary>
    /// 开发周报类
    /// 用于处理周报中项目类内容
    /// 在第一步，提取 [周报] 与 [项目计划] 比较中使用
    /// </summary>

    class DevWeeklies : BaseWeeklies
    {

        public List<WeeklyItem> CurrentWeekReleases { get; set; }
        public List<WeeklyItem> NextWeekReleasePlans { get; set; }
        public List<WeeklyItem> ProjectWorkList { get; set; }

        //区分数据分析部与开发部的每周工作
        public List<WeeklyItem> CurrentWeekWorksDataSci;
        public List<WeeklyItem> CurrentWeekWorksDevDivisions;



        public DevWeeklies() : base()
        {

            CurrentWeekReleases = new List<WeeklyItem>();
            NextWeekReleasePlans = new List<WeeklyItem>();

            ProjectWorkList = new List<WeeklyItem>();

            CurrentWeekWorksDataSci = new List<WeeklyItem>();
            CurrentWeekWorksDevDivisions = new List<WeeklyItem>();
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

            //区分分析部与开发部的工作（用于计算[分析部/开发部]*[前台/后台]共计4张表）
            CurrentWeekWorksDevDivisions = ProjectWorkList.Where(w => w.HostDivision.Contains("开发")).ToList();
            CurrentWeekWorksDataSci = ProjectWorkList.Where(w => w.HostDivision.Contains("数据")).ToList();
            

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

    }
}
