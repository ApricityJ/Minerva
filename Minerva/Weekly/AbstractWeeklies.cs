using System.Collections.Generic;


namespace Minerva.Weekly
{


    using Department;

    /// <summary>
    /// 周报集合的抽象类，用于处理周报的集合
    /// 其实没有必要有一个抽象类，我只是想试试
    /// </summary>

    abstract class AbstractWeeklies
    {

        public Dictionary<InnerDepartment, BaseWeekly> WeeklySet  { get; set; }

        //加载周报
        public abstract AbstractWeeklies Load();

        //汇总
        public abstract AbstractWeeklies Summarize();

        public abstract AbstractWeeklies Sort();

        public abstract AbstractWeeklies Save();

    }
}
