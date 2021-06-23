using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Minerva.Summary
{
    using Minerva.Weekly;
    using Minerva.Department;

    /// <summary>
    /// 周报汇总类，用于处理最终生成的周报汇总
    /// </summary>
    class Summary
    {

        public List<SummaryItem> FrontDevSummaryList { get; set; }
        public List<SummaryItem> FrontDataSummaryList { get; set; }
        public List<SummaryItem> BackDevSummaryList { get; set; }
        public List<SummaryItem> BackDataSummaryList { get; set; }



        public Summary(BaseWeekly weekly)
        {
            //FrontDevSummaryList = ToSummaryList(DepartmentType.FRONT, weekly.WeeklyDev);
            //FrontDataSummaryList = ToSummaryList(DepartmentType.FRONT, weekly.WeeklyData);
            //BackDevSummaryList = ToSummaryList(DepartmentType.BACK, weekly.WeeklyDev);
            //BackDataSummaryList = ToSummaryList(DepartmentType.BACK, weekly.WeeklyData);
        }

        private string ToRemainHumanMonth(string schedule)
        {
            return Regex.Match(schedule, @"(?<=剩余工作量：)\S+").Value;
        }
        private string ToEstimatedTimeRemaining(string schedule)
        {
            return Regex.Match(schedule, @"(?<=计划投产时间：)\S+").Value;
        }

        public SummaryItem ToSummaryItem(WeeklyItem item)
        {
            Env.Instance.SummarySequence ++;
            SummaryItem summaryItem = new SummaryItem
            {
                Sequence = Env.Instance.SummarySequence,
                ProjectName = item.Name,
                BizDepartment = item.BizDepartment,
                HostDivision = item.HostDivision,
                RemainHumanMonth = ToRemainHumanMonth(item.Schedule),
                EstimatedTimeRemaining = ToEstimatedTimeRemaining(item.Schedule)
                
            };
            return summaryItem;
        }

        private List<SummaryItem> ToSummaryList(DepartmentType type,List<WeeklyItem> weeklyItemList)
        {
            Env.Instance.SummarySequence = 0;
            return weeklyItemList
                .Where(item => Department.Instance.ToDepartmentType(item.BizDepartment).Equals(type))
                .Select(item => ToSummaryItem(item))
                .ToList();
        }
       
    }
}
