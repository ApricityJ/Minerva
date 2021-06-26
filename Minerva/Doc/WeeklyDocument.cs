using System.Collections.Generic;
using System.IO;

namespace Minerva.Doc
{
    using DAO;
    using Util;
    using Weekly;

    /// <summary>
    /// 周报文档类
    /// 用于生成 "科技与产品管理部周报#Year#年第#Week#期（#WeekEnd#）.xlsx"
    /// </summary>
    class WeeklyDocument
    {
        public string TemplatePath { get; set; }
        public string TargetPath { get; set; }

        private string template = "科技与产品管理部周报#Year#年第#Week#期（#WeekEnd#）.xlsx";

        private ExcelDAO<object> dao;
        private Dictionary<string, object> dataSet;
        private IntegratedWeeklies weeklies;

        public WeeklyDocument()
        {
            ToTemplatePath();

            dao = new ExcelDAO<object>(TemplatePath);

            ToTargetPath();
        }

        public WeeklyDocument PrepareDataSet()
        {
            weeklies = new IntegratedWeeklies();

            weeklies.Load()
                .Summarize()
                .Sort();

            dataSet = new Dictionary<string, object>();

            string title = string.Format("科技与产品管理部一周重点工作清单（{0}-{1}）",
                DateUtil.ToWeekBegin(),
                DateUtil.ToWeekEnd());

            dataSet.Add("Title", title);

            dataSet.Add("CurrentWeekWorks",
                DataTableBuilder.ToDataTableOfList(weeklies.CurrentWeekWorks, "CurrentWeekWorks"));
            dataSet.Add("CurrentWeekReleases",
                DataTableBuilder.ToDataTableOfList(weeklies.CurrentWeekReleases, "CurrentWeekReleases"));

            return this;
        }

        private void ToTemplatePath()
        {
            TemplatePath = Path.Combine("tpl", template);
        }

        private void ToTargetPath()
        {
            TargetPath = template.Replace("#Year#", DateUtil.ToCurrentYear())
                .Replace("#Week#", DateUtil.ToWeekOfYear().ToString())
                .Replace("#WeekEnd#", DateUtil.ToWeekEnd());
            TargetPath = Path.Combine(Env.Instance.RootDir, TargetPath);

            if (File.Exists(TargetPath))
            {
                File.Delete(TargetPath);
            }
        }

        public void Perform()
        {
            dao.ToDesigner()
                .SetDataSource(dataSet)
                .Process()
                .SaveAs(TargetPath);
        }

    }
}
