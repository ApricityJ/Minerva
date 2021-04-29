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
            WorkReportDivision1st = new ExcelDesolator<WorkReportItem>(Env.Instance.ReportDivsion1st)
                .ToEntityList(0, typeof(WorkReportItem));
            WorkReportDivision2nd = new ExcelDesolator<WorkReportItem>(Env.Instance.ReportDivsion2nd)
                .ToEntityList(0, typeof(WorkReportItem));
            WorkReportDivision3rd = new ExcelDesolator<WorkReportItem>(Env.Instance.ReportDivsion3rd)
                .ToEntityList(0, typeof(WorkReportItem));
            WorkReportDivisionData = new ExcelDesolator<WorkReportItem>(Env.Instance.ReportDivsionData)
                .ToEntityList(0, typeof(WorkReportItem));
            WorkReportList = WorkReportDivision1st
                .Concat(WorkReportDivision2nd)
                .Concat(WorkReportDivision3rd)
                .Concat(WorkReportDivisionData)
                .ToList();
            return this;
        }

        private List<WorkReportItem> ToWorkReport(string path)
        {
            List<WorkReportItem> workReportList = new ExcelDesolator<WorkReportItem>(path)
                .ToEntityList(0, typeof(WorkReportItem));

            //workReportList = workReportList.Where()

            return workReportList;
        }






    }
}
