using System.Collections.Generic;


namespace Minerva.Weekly
{

    /// <summary>
    /// 开发部门周报
    /// 增加了[本周投产情况]与[下周投产计划]部分
    /// </summary>
    class DevWeekly :BaseWeekly
    {
        public List<WeeklyItem> CurrentWeekRelease { get; set; }
        public List<WeeklyItem> NextWeekReleasePlan { get; set; }

        public DevWeekly(string path) : base(path)
        {
            CurrentWeekRelease = ExtractBy("本周投产情况");
            NextWeekReleasePlan = ExtractBy("下周投产计划");
        }

    }
}
