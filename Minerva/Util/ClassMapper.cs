using System;
using System.Collections.Generic;



namespace Minerva.Util
{
    using Minerva.Weekly;
    using Minerva.Project;

    /// <summary>
    /// 
    /// </summary>
    class ClassMapper
    {
        public static readonly Dictionary<Type, Dictionary<int,string>> ClassMap = new Dictionary<Type, Dictionary<int, string>>();


        static ClassMapper()
        {
            Dictionary<int, string> WeeklyItemMap = new Dictionary<int, string>
            {
                { 0, "Sequence" },
                { 1, "Name" },
                { 2, "Type" },
                { 3, "CurrentProgress" },
                { 4, "HostDivision" },
                { 5, "CoHostDivision" },
                { 6, "BizDepartment" },
                { 7, "ResponsiblePersonnel" },
                { 8, "Participants" },
                { 9, "Schedule" }
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
