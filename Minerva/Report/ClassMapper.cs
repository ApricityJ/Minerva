using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace Minerva.Report
{
    class ClassMapper
    {

        private static readonly ClassMapper instance = new ClassMapper();
        public static Dictionary<Type, Dictionary<int,string>> classMap = new Dictionary<Type, Dictionary<int, string>>();


        static ClassMapper()
        {
            Dictionary<int, string> WorkReportItemMap = new Dictionary<int, string>
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

            classMap.Add(typeof(Report.WorkReportItem), WorkReportItemMap);

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

            classMap.Add(typeof(Report.ProjectPlanItem), ProjectPlanItemMap);

        }

        private ClassMapper()
        {
        }

        public static ClassMapper Instance
        {
            get
            {
                return instance;
            }
        }

        public Dictionary<int, string> ToClassMap(Type type)
        {
            return classMap[type];
        }





    }
}
