using System.Collections.Generic;
using System.Linq;

namespace Minerva.Weekly
{
    using Minerva.Department;
    using Minerva.DAO;
    using System.IO;
    using Minerva.Util;

    /// <summary>
    /// 周报类，用于处理周报的集合
    /// </summary>
    internal class Weekly
    {
        public List<WeeklyItem> WeeklyDivision1st { get; set; }
        public List<WeeklyItem> WeeklyDivision2nd { get; set; }
        public List<WeeklyItem> WeeklyDivision3rd { get; set; }
        public List<WeeklyItem> WeeklyData { get; set; }
        public List<WeeklyItem> WeeklyList { get; set; }
        public List<WeeklyItem> WeeklyDev { get; set; }
        public List<WeeklyItem> WeeklySorted { get; set; }

        private string Template = "科技与产品管理部周报(#Year#)年第(#Week#)期.et";
        private string TargetPath { get; set; }

        public Weekly()
        {
            WeeklyDivision1st = ToWeekly(Env.Instance.ReportDivsion1st);
            WeeklyDivision2nd = ToWeekly(Env.Instance.ReportDivsion2nd);
            WeeklyDivision3rd = ToWeekly(Env.Instance.ReportDivsion3rd);
            WeeklyData = ToWeekly(Env.Instance.ReportDivsionData);

            WeeklyDev = WeeklyDivision1st
                .Concat(WeeklyDivision2nd)
                .Concat(WeeklyDivision3rd)
                .ToList();

            WeeklyList = WeeklyDev
                .Concat(WeeklyData)
                .ToList();

        }


        //判断是否是项目类工作


        //获取一个部门的项目周报(仅限开发和数据分析部)
        private List<WeeklyItem> ToWeekly(string path)
        {
            
            if (path.Equals(""))
            {
                return new List<WeeklyItem>();
            }

            //读取excel文件
            List<WeeklyItem> weeklyItemList = new ExcelDAO<WeeklyItem>(path)
                .SelectSheetAt(0)
                .Skip(2)
                .ToEntityList(typeof(WeeklyItem));

            //尝试将业务部门的名称转为正式名称
            weeklyItemList.ForEach(item =>
            {
                item.BizDepartment = Department.Instance.ToDepartmentName(item.BizDepartment);
            });

            //仅保留项目类工作，仅考虑已立项项目(不包含立项和需求中的项目)
            return weeklyItemList
                .Where(item => item.IsProjectWork())
                //.Where(item => !item.IsProjectApproved())
                .ToList();

        }


        public Weekly ToSortedWeeklyList()
        {
            WeeklyDev.Sort();
            WeeklyData.Sort();
            
            WeeklySorted = WeeklyDev.Concat(WeeklyData).ToList();

            return this;
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

        public Weekly Summarize()
        {
            ToTargetPath();

            ExcelDAO<WeeklyItem> dao = new ExcelDAO<WeeklyItem>(TargetPath);
            dao.SetCellValues(this.WeeklySorted);
            dao.Save();

            return this;
        }

        public bool IsExistInWeeklyList(string projectName)
        {
            return WeeklyList.Any(item => item.Name.Trim().Equals(projectName));

        }


    }
}
