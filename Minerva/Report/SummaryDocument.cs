using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;

namespace Minerva.Report
{
    class SummaryDocument
    {
        private string TemplatePath { get; set; }
        private string TargetPath { get; set; }
        private Dictionary<string, object> FieldDataSet { get; set; }

        private string TemplateBaseDir = "tpl";
        private string Template = "科技与产品管理部周报(#Year#)年第(#Week#)期.docx";
        private List<DataTable> DataTableList;

        private Summary summary;


        public SummaryDocument(Summary summary)
        {
            this.summary = summary;

            DataTableList = new List<DataTable>();
            DataTableList.Add(DataTableUtil.ToDataTableOfList(summary.FrontDevSummaryList, "FrontDevSummaryList"));
            DataTableList.Add(DataTableUtil.ToDataTableOfList(summary.FrontDataSummaryList, "FrontDataSummaryList"));
            DataTableList.Add(DataTableUtil.ToDataTableOfList(summary.BackDevSummaryList, "BackDevSummaryList"));
            DataTableList.Add(DataTableUtil.ToDataTableOfList(summary.BackDataSummaryList, "BackDataSummaryList"));

            ToTemplatePath();
            ToTargetPath();
        }

        private void ToTemplatePath()
        {
            TemplatePath = Path.Combine(TemplateBaseDir, Template);
        }

        private int ToWeekOfYear()
        {
            CultureInfo ci = new CultureInfo("zh-CN");
            System.Globalization.Calendar cal = ci.Calendar;
            CalendarWeekRule cwr = ci.DateTimeFormat.CalendarWeekRule;
            DayOfWeek dow = DayOfWeek.Monday;
            return cal.GetWeekOfYear(DateTime.Now, cwr, dow);
        }

        private void ToTargetPath()
        {
            TargetPath = Template.Replace("#Year#", DateTime.Now.Year.ToString());
            TargetPath = TargetPath.Replace("#Week#", ToWeekOfYear().ToString());
        }


        public void Perform()
        {
            Document doc = new Document(TemplatePath);

            List<string> names = new List<string>();
            List<object> values = new List<object>();


            //FieldDataSet = new Dictionary<string, object>();
            //FieldDataSet.ToList().ForEach(pair =>
            //{
            //    names.Add(pair.Key);
            //    values.Add(pair.Value);
            //});
            //doc.MailMerge.Execute(names.ToArray(), values.ToArray());

            DataTableList.ForEach(table =>
            {
                doc.MailMerge.ExecuteWithRegions(table);
            });

            doc.Save(TargetPath, SaveFormat.Docx);
        }


    }
}
