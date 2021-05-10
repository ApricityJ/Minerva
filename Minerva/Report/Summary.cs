using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Minerva.Report
{
    class Summary
    {

        public List<SummaryItem> FrontDevSummaryList { get; set; }
        public List<SummaryItem> FrontDataSummaryList { get; set; }
        public List<SummaryItem> BackDevSummaryList { get; set; }
        public List<SummaryItem> BackDataSummaryList { get; set; }

        public Summary()
        {

        }

        public Summary(WorkReport report)
        {
            FrontDevSummaryList = ToSummaryList(DepartmentType.FRONT, report.WorkReportDev);
            FrontDataSummaryList = ToSummaryList(DepartmentType.FRONT, report.WorkReportData);
            BackDevSummaryList = ToSummaryList(DepartmentType.BACK, report.WorkReportDev);
            BackDataSummaryList = ToSummaryList(DepartmentType.BACK, report.WorkReportData);
        }

        public string ToRemainHumanMonth(string schedule)
        {
            return Regex.Match(schedule, @"(?<=剩余工作量：)\S+").Value;
        }
        public string ToEstimatedTimeRemaining(string schedule)
        {
            return Regex.Match(schedule, @"(?<=计划投产时间：)\S+").Value;
        }

        public SummaryItem ToSummaryItem(WorkReportItem workReportItem)
        {
            Env.Instance.SummarySequence ++;
            SummaryItem summaryItem = new SummaryItem
            {
                Sequence = Env.Instance.SummarySequence,
                ProjectName = workReportItem.Name,
                BizDepartment = workReportItem.BizDepartment,
                HostDivision = workReportItem.HostDivision,
                RemainHumanMonth = ToRemainHumanMonth(workReportItem.Schedule),
                EstimatedTimeRemaining = ToEstimatedTimeRemaining(workReportItem.Schedule)
                
            };
            return summaryItem;
        }

        private List<SummaryItem> ToSummaryList(DepartmentType type,List<WorkReportItem> workReportList)
        {
            Env.Instance.SummarySequence = 0;
            return workReportList
                .Where(reportItem => Department.Instance.ToDepartmentType(reportItem.BizDepartment).Equals(type))
                .Select(reportItem=>ToSummaryItem(reportItem))
                .ToList();
        }
       
    }
}
