using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Minerva.Weekly
{
    using DAO;
    using Department;
    using Util;

    /// <summary>
    /// 周报集成类，包含一个开发类周报集合和一个基础类周报集合
    /// 用于生成 [科技与产品管理部周报(#Year#)年第(#Week#)期.et]
    /// </summary>

    class IntegratedWeeklies : DevWeeklies
    {

        public DevWeeklies DevelopWeeklies { get; set; }

        public BaseWeeklies NonDevelopWeeklies { get; set; }


        public IntegratedWeeklies()
        {
            WeeklySet = new Dictionary<InnerDepartment, BaseWeekly>();

            DevelopWeeklies = new DevWeeklies();
            NonDevelopWeeklies = new BaseWeeklies();

            CurrentWeekWorks = new List<WeeklyItem>();
            UnnormalCases = new List<WeeklyItem>();

        }


        public override AbstractWeeklies Load()
        {
            DevelopWeeklies.Load();
            NonDevelopWeeklies.Load();

            return this;
        }

        public override AbstractWeeklies Summarize()
        {
            
            DevelopWeeklies.Summarize();

            CurrentWeekReleases = DevelopWeeklies.CurrentWeekReleases;
            NextWeekReleasePlans = DevelopWeeklies.NextWeekReleasePlans;


            NonDevelopWeeklies.Summarize();

            CurrentWeekWorks = CurrentWeekWorks.Concat(DevelopWeeklies.CurrentWeekWorks)
                .Concat(NonDevelopWeeklies.CurrentWeekWorks).ToList();
            UnnormalCases = UnnormalCases.Concat(DevelopWeeklies.UnnormalCases)
                .Concat(NonDevelopWeeklies.UnnormalCases).ToList();

            return this;
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
