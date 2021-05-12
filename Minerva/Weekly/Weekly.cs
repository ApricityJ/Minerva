using System.Collections.Generic;
using System.Linq;

namespace Minerva.Weekly
{
    using Minerva.Department;
    using Minerva.DAO;

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
        private bool IsProjectWork(WeeklyItem item)
        {
            return item.Type.ToString().Contains("项目");
        }

        //获取一个部门的项目周报(仅限开发和数据分析部)
        private List<WeeklyItem> ToWeekly(string path)
        {
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

            //仅保留项目类工作
            return weeklyItemList
                .Where(item => IsProjectWork(item))
                .ToList();


        }






    }
}
