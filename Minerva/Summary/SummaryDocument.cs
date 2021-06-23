using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace Minerva.Summary
{

    using Minerva.Util;
    /// <summary>
    /// 汇总 -- 科技与产品管理部周报(#Year#)年第(#Week#)期
    /// 用于生成上述文档，请注意参考模板
    /// </summary>
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
            DataTableList.Add(DataTableBuilder.ToDataTableOfList(summary.FrontDevSummaryList, "FrontDevSummaryList"));
            DataTableList.Add(DataTableBuilder.ToDataTableOfList(summary.FrontDataSummaryList, "FrontDataSummaryList"));
            DataTableList.Add(DataTableBuilder.ToDataTableOfList(summary.BackDevSummaryList, "BackDevSummaryList"));
            DataTableList.Add(DataTableBuilder.ToDataTableOfList(summary.BackDataSummaryList, "BackDataSummaryList"));

            ToTemplatePath();
            ToTargetPath();
        }

        private void ToTemplatePath()
        {
            TemplatePath = Path.Combine(TemplateBaseDir, Template);
        }

        //计算当前为一年的第n周


        private void ToTargetPath()
        {
            TargetPath = Template.Replace("#Year#", DateUtil.ToCurrentYear())
                .Replace("#Week#", DateUtil.ToWeekOfYear().ToString());
            TargetPath = Path.Combine(Env.Instance.RootDir, TargetPath);

            if (File.Exists(TargetPath))
            {
                File.Delete(TargetPath);
            }
        }


        public void Perform()
        {
            Document doc = new Document(TemplatePath);

            //如果需要替换单个变量，请参考如下代码：

            //List<string> names = new List<string>();
            //List<object> values = new List<object>();

            //FieldDataSet = new Dictionary<string, object>();
            //FieldDataSet.Add("key","value");
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
