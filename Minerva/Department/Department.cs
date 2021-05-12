using System.Collections.Generic;
using System.Linq;

namespace Minerva.Department
{
    /// <summary>
    /// 部门类，用于管理用到的部门及模糊处理周报中的部门名称
    /// 全局仅需要一个，使用单例
    /// </summary>
    public class Department
    {
        private readonly Dictionary<string, DepartmentType> departmentMap;

        private Department()
        {
            departmentMap = new Dictionary<string, DepartmentType>()
            {
                {"公司业务部",DepartmentType.FRONT },
                {"房产金融部",DepartmentType.FRONT},
                {"房地产金融部",DepartmentType.FRONT},
                {"金融同业部",DepartmentType.FRONT},
                {"国际金融部",DepartmentType.FRONT},
                {"信用卡中心",DepartmentType.FRONT},
                {"个人金融部",DepartmentType.FRONT},
                {"个人信贷部",DepartmentType.FRONT},
                {"机构业务部",DepartmentType.FRONT},
                {"普惠金融部",DepartmentType.FRONT},
                {"网络金融部",DepartmentType.FRONT},
                {"资产负债部",DepartmentType.FRONT},
                {"运营管理部",DepartmentType.BACK},
                {"办公室",DepartmentType.BACK},
                {"内控合规部",DepartmentType.BACK},
                {"科技与产品管理部",DepartmentType.BACK},
                {"团委",DepartmentType.BACK},
                {"工会",DepartmentType.BACK},
                {"总务部",DepartmentType.BACK}

            };
        }

        public static Department Instance { get; } = new Department();

        //根据部门类型获取对应的部门名称List
        public List<string> ToDepartmentList(DepartmentType type)
        {
            return departmentMap
                .Where(pair => pair.Value.Equals(type))
                .Select(pair => pair.Key)
                .ToList();
        }

        //判断周报中的所填写的部门名称是不是能否与正式部门名称完整匹配
        //e.g.
        //IsCompletelyMatch("个人金融部","个人金融部") -> true
        //IsCompletelyMatch("个人金融部","个金") -> true
        //IsCompletelyMatch("个人金融部","个人业务部") -> false
        private bool IsCompletelyMatch(string department,string nameToMatch)
        {
            return nameToMatch.All(c => department.Contains(c));
        }

        //将周报中的部门名称转为正式部门名称，如果找不到则使用周报中的部门名称
        //e.g.
        //ToDepartmentName("个金") -> 个人金融部
        //ToDepartmentName("个人金融部") -> 个人金融部
        //ToDepartmentName("个人业务部") -> 个人业务部
        public string ToDepartmentName(string department)
        {
            List<string> list = departmentMap.Where(pair => IsCompletelyMatch(pair.Key, department))
                .Select(pair => pair.Key)
                .ToList();

            if (list.Count > 0)
            {
                return list[0];
            }

            return department;
        }

        //根据部门名称返回部门类型
        //e.g. 
        //ToDepartmentType("个人金融部") -> FRONT
        //ToDepartmentType("工会") -> BACK
        //ToDepartmentType("个人业务部") -> UNKNOWN
        public DepartmentType ToDepartmentType(string department)
        {
            string dept = ToDepartmentName(department);

            if (departmentMap.ContainsKey(dept))
            {
                return departmentMap[dept];
            }

            return DepartmentType.UNKNOWN;
        }


    }
}
