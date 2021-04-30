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

        private List<SummaryItem> frontDevSummaryList;
        private List<SummaryItem> frontDataSummaryList;
        private List<SummaryItem> backDevSummaryList;
        private List<SummaryItem> backDataSummaryList;


        public Summary(WorkReport report)
        {
            frontDevSummaryList = ToSummaryList(DepartmentType.FRONT, report.WorkReportDev);
            frontDataSummaryList = ToSummaryList(DepartmentType.FRONT, report.WorkReportData);
            backDevSummaryList = ToSummaryList(DepartmentType.BACK, report.WorkReportDev);
            backDataSummaryList = ToSummaryList(DepartmentType.BACK, report.WorkReportData);
        }

        private string ToRemainHumanMonth(string schedule)
        {
            return Regex.Match(schedule, @"(?<=剩余工作量：)\S+").Value;
        }
        private string ToEstimatedTimeRemaining(string schedule)
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
