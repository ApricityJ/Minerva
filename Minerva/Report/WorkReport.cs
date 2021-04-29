using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Report
{
    class WorkReport
    {
        public List<WorkReportItem> WorkReportDivision1st { get; set; }
        public List<WorkReportItem> WorkReportDivision2nd { get; set; }
        public List<WorkReportItem> WorkReportDivision3rd { get; set; }
        public List<WorkReportItem> WorkReportDivisionData { get; set; }
        public List<WorkReportItem> WorkReportList { get; set; }




        public WorkReport()
        {

        }


        public WorkReport Initalize()
        {
            WorkReportDivision1st = ToWorkReport(Env.Instance.ReportDivsion1st);
            WorkReportDivision2nd = ToWorkReport(Env.Instance.ReportDivsion2nd);
            WorkReportDivision3rd = ToWorkReport(Env.Instance.ReportDivsion3rd);
            WorkReportDivisionData = ToWorkReport(Env.Instance.ReportDivsionData);
            WorkReportList = WorkReportDivision1st
                .Concat(WorkReportDivision2nd)
                .Concat(WorkReportDivision3rd)
                .Concat(WorkReportDivisionData)
                .ToList();
            return this;
        }

        private bool IsProjectWork(WorkReportItem item)
        {
            return item.Type.ToString().Contains("项目");
        }

        private List<WorkReportItem> ToWorkReport(string path)
        {
            List<WorkReportItem> workReportList = new ExcelDesolator<WorkReportItem>(path)
                .ToEntityList(0, 2, typeof(WorkReportItem));

            return workReportList
                .Where(workReportItem => IsProjectWork(workReportItem))
                .ToList();
        }






    }
}
