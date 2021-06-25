using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Minerva.Brief
{
    using Minerva.Weekly;
    using Minerva.Department;

    /// <summary>
    /// 开发工作每周概况类，用于处理最终开发工作每周概况
    /// </summary>
    class DevBrief
    {

        public List<BriefItem> FrontDevBriefList { get; set; }
        public List<BriefItem> FrontDataBriefList { get; set; }
        public List<BriefItem> BackDevBriefList { get; set; }
        public List<BriefItem> BackDataBriefList { get; set; }



        public DevBrief()
        {
            DevWeeklies weeklies = new DevWeeklies();

            weeklies.Load().Summarize();

            FrontDevBriefList = ToBriefItemList(DepartmentType.FRONT, weeklies.CurrentWeekWorksDevDivisions);
            FrontDataBriefList = ToBriefItemList(DepartmentType.FRONT, weeklies.CurrentWeekWorksDataSci);
            BackDevBriefList = ToBriefItemList(DepartmentType.BACK, weeklies.CurrentWeekWorksDevDivisions);
            BackDataBriefList = ToBriefItemList(DepartmentType.BACK, weeklies.CurrentWeekWorksDataSci);
        }

        private string ToRemainHumanMonth(string schedule)
        {
            return Regex.Match(schedule, @"(?<=剩余工作量：)\S+").Value;
        }
        private string ToEstimatedTimeRemaining(string schedule)
        {
            return Regex.Match(schedule, @"(?<=计划投产时间：)\S+").Value;
        }

        private BriefItem ToBriefItem(WeeklyItem item)
        {
            Env.Instance.BriefSequence++;
            BriefItem briefItem = new BriefItem
            {
                Sequence = Env.Instance.BriefSequence,
                ProjectName = item.Name,
                BizDepartment = item.BizDepartment,
                HostDivision = item.HostDivision,
                RemainHumanMonth = ToRemainHumanMonth(item.Schedule),
                EstimatedTimeRemaining = ToEstimatedTimeRemaining(item.Schedule)
                
            };
            return briefItem;
        }

        private List<BriefItem> ToBriefItemList(DepartmentType type,List<WeeklyItem> weeklyItemList)
        {
            Env.Instance.BriefSequence = 0;
            return weeklyItemList
                .Where(item => Department.Instance.ToDepartmentType(item.BizDepartment).Equals(type))
                .Select(item => ToBriefItem(item))
                .ToList();
        }
       
    }
}
