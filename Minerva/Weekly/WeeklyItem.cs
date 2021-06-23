using System;
using System.Collections.Generic;
using System.Linq;

namespace Minerva.Weekly
{
    public class WeeklyItem : IComparer<WeeklyItem>, IComparable<WeeklyItem>
    {
        static Dictionary<string, int> DivisionWeightMap = new Dictionary<string, int>()
        {
            { "一部", 1 },
            { "二部", 2 },
            { "三部", 3 },
            { "数据", 4 },
            { "测试", 5 },
            { "系统", 6 },
            { "网络", 7 },
            { "产品", 8 }
        };

        static Dictionary<string, int> ProjectTypeWeightMap = new Dictionary<string, int>()
        {

            { "一般", 1 },
            { "快捷", 2 },
            { "数据", 3 },
            {"重要需求",4 },
            {"一般需求",5 },
            {"其它",6 }

        };



        //序号	项目/需求/管理名称	任务类型   项目编号（当前状态）	当前进展	主办部门	协办部门	业务部门	负责人员	参与人员	进度计划
        public int Sequence { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string ProjectId { get; set; }
        public string CurrentProgress { get; set; }
        public string HostDivision { get; set; }
        public string CoHostDivision { get; set; }
        public string BizDepartment { get; set; }
        public string ResponsiblePersonnel { get; set; }
        public string Participants { get; set; }
        public string Schedule { get; set; }

        public int ToDivisionWeight()
        {
            return DivisionWeightMap.FirstOrDefault(pair => HostDivision.Contains(pair.Key)).Value;
        }

        public int ToProjectTypeWeight()
        {
            return ProjectTypeWeightMap.FirstOrDefault(pair => Type.Contains(pair.Key)).Value;
        }

        public bool IsProjectWork()
        {
            return Type.Contains("项目");
        }

        public bool IsSuspended()
        {
            string status = CurrentProgress.Split('\n')[0];
            return status.Contains("暂停");
        }

        public bool IsProjectApproved()
        {
            string status = CurrentProgress.Split('\n')[0];
            return status.Contains("立项") || status.Contains("需求");
        }

        public int Compare(WeeklyItem x, WeeklyItem y)
        {
            int diff = x.ToProjectTypeWeight() - y.ToProjectTypeWeight();
            if (diff == 0)
            {
                return x.ToDivisionWeight() - y.ToDivisionWeight();
            }
            return diff;
        }

        public int CompareTo(WeeklyItem other)
        {
            return Compare(this, other);
        }
    }
}
