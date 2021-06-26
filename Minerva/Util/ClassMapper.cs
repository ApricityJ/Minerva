using System;
using System.Collections.Generic;



namespace Minerva.Util
{
    using Minerva.Weekly;
    using Minerva.Project;

    /// <summary>
    /// 类的映射管理
    /// 用于从Excel的[行]到一个类的[实例]的对照关系转换
    /// 用于读取[项目计划]及[工作周报]内容
    /// </summary>
    class ClassMapper
    {
        public static readonly Dictionary<Type, Dictionary<int, string>> ClassMap = new Dictionary<Type, Dictionary<int, string>>();


        static ClassMapper()
        {
            Dictionary<int, string> WeeklyItemMap = new Dictionary<int, string>
            {
                { 0, "Sequence" },
                { 1, "Name" },
                { 2, "Type" },
                { 3, "ProjectId" },
                { 4, "CurrentProgress" },
                { 5, "HostDivision" },
                { 6, "CoHostDivision" },
                { 7, "BizDepartment" },
                { 8, "ResponsiblePersonnel" },
                { 9, "Participants" },
                { 10, "Schedule" }
            };

            ClassMap.Add(typeof(WeeklyItem), WeeklyItemMap);

            Dictionary<int, string> ProjectPlanItemMap = new Dictionary<int, string>
            {
                { 0, "Sequence" },
                { 1, "ProjectName" },
                { 2, "RequirementDepartment" },
                { 3, "InnovationType" },
                { 4, "Breif" },
                { 5, "YearOfApproval" },
                { 6, "DateOfRelease" },
                { 7, "ProjectType" },
                { 8, "Budget" },
                { 9, "IsDigitalTransformation" },
                { 10, "Status" },
                { 11, "HostDivision" },
                { 12, "ResponsiblePersonnel" },
                { 13, "MotivationalSuggestion" }
            };

            ClassMap.Add(typeof(ProjectPlanItem), ProjectPlanItemMap);

        }

        private ClassMapper()
        {
        }

        public static ClassMapper Instance { get; } = new ClassMapper();

        public Dictionary<int, string> ToClassMap(Type type)
        {
            return ClassMap[type];
        }





    }
}
