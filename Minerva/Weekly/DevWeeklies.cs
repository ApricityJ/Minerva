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

    class DevWeeklies : AbstractWeelies
    {

        public List<WeeklyItem> CurrentWeekReleases { get; set; }
        public List<WeeklyItem> NextWeekReleasePlans { get; set; }

        public List<DevWeekly> DevWeeklyList { get; set; }
        public List<WeeklyItem> ProjectWorkList { get; set; }


        public DevWeeklies()
        {
            DevWeeklyList = new List<DevWeekly>();
            ProjectWorkList = new List<WeeklyItem>();
        }


        //加载周报
        public override AbstractWeelies LoadWeekly()
        {
            foreach (InnerDepartment department in Enum.GetValues(typeof(InnerDepartment)))
            {
                if (department.ToString().StartsWith("Dev"))
                {
                    WeeklyCollection.Add(department, new BaseWeekly(Env.Instance.ToWeeklyPath(department)));
                }
            }
            return this;
        }

        //筛选项目相关的工作内容
        public override AbstractWeelies Summarize()
        {
            //把1，2，3，4部的本周工作报告拼在一起
            DevWeeklyList.ForEach(w => ProjectWorkList.Concat(w.CurrentWeekWork));

            //筛选项目类工作（用于与工作计划交互匹配）
            ProjectWorkList = ProjectWorkList.Where(item => item.IsProjectWork())
                .ToList();

            return this;
        }

        //判断一个项目名称是否出现在周报中
        //用于寻找 [项目计划] 中存在 但 [周报] 中不存在的项目
        public bool IsExist(string projectName)
        {
            return ProjectWorkList.Any(item => item.Name.Trim().Equals(projectName.Trim()));

        }

    }
}
