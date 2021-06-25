using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Minerva.Doc
{

    using Util;
    using DAO;
    using Brief;
    /// <summary>
    /// 汇总 -- 开发工作每周概况（#WeekBegin#-#WeekEnd#）
    /// 用于生成上述文档，请注意参考模板
    /// </summary>
    class DevBriefDocument
    {
        public string TemplatePath { get; set; }
        public string TargetPath { get; set; }
        public Dictionary<string, object> FieldSet { get; set; }
        public Dictionary<string, DataTable> TableSet { get; set; }

        private string template = "开发工作每周概况（#WeekBegin#-#WeekEnd#）.docx";
        private DevBrief brief;

        public DevBriefDocument()
        {
            ToTemplatePath();

            brief = new DevBrief();

            ToTargetPath();
        }


        public DevBriefDocument PrepareDataSet()
        {
            ToTableSet();

            ToFieldSet();

            return this;
        }

        private void ToTableSet()
        {
            TableSet = new Dictionary<string, DataTable>();
            TableSet.Add("FrontDevBriefList",
                DataTableBuilder.ToDataTableOfList(brief.FrontDevBriefList, "FrontDevBriefList"));
            TableSet.Add("FrontDataBriefList",
                DataTableBuilder.ToDataTableOfList(brief.FrontDataBriefList, "FrontDataBriefList"));
            TableSet.Add("BackDevBriefList",
                DataTableBuilder.ToDataTableOfList(brief.BackDevBriefList, "BackDevBriefList"));
            TableSet.Add("BackDataBriefList",
                DataTableBuilder.ToDataTableOfList(brief.BackDataBriefList, "BackDataBriefList"));
        }

        private void ToFieldSet()
        {
            FieldSet = new Dictionary<string, object>();
            FieldSet.Add("WeekBegin", DateUtil.ToWeekBegin());
            FieldSet.Add("WeekEnd", DateUtil.ToWeekEnd());
        }

        private void ToTemplatePath()
        {
            TemplatePath = Path.Combine("tpl", template);
        }


        private void ToTargetPath()
        {
            TargetPath = template.Replace("#WeekBegin#", DateUtil.ToWeekBegin())
                .Replace("#WeekEnd#", DateUtil.ToWeekEnd());
            TargetPath = Path.Combine(Env.Instance.RootDir, TargetPath);

            if (File.Exists(TargetPath))
            {
                File.Delete(TargetPath);
            }
        }


        public void Perform()
        {
            new WordDAO(TemplatePath)
                .SetFieldSet(FieldSet)
                .SetTableSet(TableSet)
                .Process()
                .SaveAs(TargetPath);
        }
    }
}
