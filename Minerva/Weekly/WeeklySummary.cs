using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Weekly
{
    using Minerva.Department;
    using Minerva.DAO;
    using System.IO;
    using Minerva.Util;
    class WeeklySummary
    {
        AbstractWeekly WeeklyDevDivision1st = new AbstractWeekly(Env.Instance.ReportDivsion1st);
        AbstractWeekly WeeklyDevDivision2nd = new AbstractWeekly(Env.Instance.ReportDivsion2nd);
        AbstractWeekly WeeklyDevDivision3rd = new AbstractWeekly(Env.Instance.ReportDivsion3rd);
        AbstractWeekly WeeklyDataSci = new AbstractWeekly(Env.Instance.ReportDivsionData);
        AbstractWeekly WeeklyTest = new AbstractWeekly(Env.Instance.ReportTest);
        AbstractWeekly WeeklySystem = new AbstractWeekly(Env.Instance.ReportSystem);
        AbstractWeekly WeeklyNet = new AbstractWeekly(Env.Instance.ReportNet);
        AbstractWeekly WeeklyProduct = new AbstractWeekly(Env.Instance.ReportProduct);

        Dictionary<string, AbstractWeekly> WeeklyCollection = new Dictionary<string, AbstractWeekly>();

        List<WeeklyItem> SummaryReport;
        List<WeeklyItem> CurrentWeekReleases = new List<WeeklyItem>();
        List<WeeklyItem> NextWeekReleases = new List<WeeklyItem>();
        List<WeeklyItem> Exceptions = new List<WeeklyItem>();

        private string Template = "科技与产品管理部周报(#Year#)年第(#Week#)期.et";
        private string TargetPath { get; set; }


        public WeeklySummary()
        {
            WeeklyCollection.Add("Division1st", WeeklyDevDivision1st);
            WeeklyCollection.Add("Division2nd", WeeklyDevDivision2nd);
            WeeklyCollection.Add("Division3rd", WeeklyDevDivision3rd);
            WeeklyCollection.Add("DataSci", WeeklyDataSci);
            WeeklyCollection.Add("Test", WeeklyTest);
            WeeklyCollection.Add("System", WeeklySystem);
            WeeklyCollection.Add("Net", WeeklyNet);
            WeeklyCollection.Add("Product", WeeklyProduct);

            SummaryReport = WeeklyDevDivision1st.Report
                .Concat(WeeklyDevDivision2nd.Report)
                .Concat(WeeklyDevDivision3rd.Report)
                .Concat(WeeklyDataSci.Report)
                .ToList();

            SummaryReport.Sort();
            SummaryReport = SummaryReport.
                Concat(WeeklyTest.Report)
                .Concat(WeeklySystem.Report)
                .Concat(WeeklyNet.Report)
                .Concat(WeeklyProduct.Report).ToList();

            WeeklyCollection.Select(w => CurrentWeekReleases.Concat(w.Value.CurrentWeekRelease));
            WeeklyCollection.Select(w => NextWeekReleases.Concat(w.Value.NextWeekRelease));
            WeeklyCollection.Select(w =>Exceptions.Concat(w.Value.ExceptionReport));

            CurrentWeekReleases.Sort();
            NextWeekReleases.Sort();

        }


        private void ToTargetPath()
        {
            TargetPath = Template.Replace("#Year#", DateUtil.ToCurrentYear())
                .Replace("#Week#", DateUtil.ToWeekOfYear().ToString());
            TargetPath = Path.Combine(Env.Instance.WeeklyReportsDir, TargetPath);

            if (File.Exists(TargetPath))
            {
                File.Delete(TargetPath);
            }
        }

        public WeeklySummary Summarize()
        {
            ToTargetPath();

            ExcelDAO<WeeklyItem> dao = new ExcelDAO<WeeklyItem>(TargetPath);
            dao.SetCellValues(this.SummaryReport);
            dao.Save();

            return this;
        }




    }
}
