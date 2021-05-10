using System.Collections.Generic;
using System.Linq;

namespace Minerva.Report
{
    internal class Department
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

        public List<string> ToDepartmentList(DepartmentType type)
        {
            return departmentMap
                .Where(pair => pair.Value.Equals(type))
                .Select(pair => pair.Key)
                .ToList();
        }

        private bool IsCompletelyMatch(string department,string nameToMatch)
        {
            return nameToMatch.All(c => department.Contains(c));
        }

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
