using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Report
{
    internal class WorkReport
    {
        public List<WorkReportItem> WorkReportDivision1st { get; set; }
        public List<WorkReportItem> WorkReportDivision2nd { get; set; }
        public List<WorkReportItem> WorkReportDivision3rd { get; set; }
        public List<WorkReportItem> WorkReportData { get; set; }
        public List<WorkReportItem> WorkReportList { get; set; }
        public List<WorkReportItem> WorkReportDev { get; set; }


        public WorkReport()
        {
            WorkReportDivision1st = ToWorkReport(Env.Instance.ReportDivsion1st);
            WorkReportDivision2nd = ToWorkReport(Env.Instance.ReportDivsion2nd);
            WorkReportDivision3rd = ToWorkReport(Env.Instance.ReportDivsion3rd);
            WorkReportData = ToWorkReport(Env.Instance.ReportDivsionData);

            WorkReportDev = WorkReportDivision1st
                .Concat(WorkReportDivision2nd)
                .Concat(WorkReportDivision3rd)
                .ToList();

            WorkReportList = WorkReportDev
                .Concat(WorkReportData)
                .ToList();
        }



        private bool IsProjectWork(WorkReportItem item)
        {
            return item.Type.ToString().Contains("项目");
        }

        private List<WorkReportItem> ToWorkReport(string path)
        {
            List<WorkReportItem> workReportList = new ExcelDesolator<WorkReportItem>(path)
                .SelectSheetAt(0)
                .Skip(2)
                .ToEntityList(typeof(WorkReportItem));

            workReportList.ForEach(workReportItem =>
            {
                workReportItem.BizDepartment = Department.Instance.ToDepartmentName(workReportItem.BizDepartment);
            });

            return workReportList
                .Where(workReportItem => IsProjectWork(workReportItem))
                .ToList();


        }






    }
}
