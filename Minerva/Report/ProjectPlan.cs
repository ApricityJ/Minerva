using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Report
{
    

    class ProjectPlan
    {
        private ExcelDesolator<ProjectPlanItem> desolator;
        private List<ProjectPlanItem> projectPlanList;
        private List<ProjectPlanItem> ignoredProjectList;

        public ProjectPlan()
        {
            desolator = new ExcelDesolator<ProjectPlanItem>(Env.Instance.ProjectPlanPath);
            projectPlanList = desolator.ToEntityList(0, 2, typeof(ProjectPlanItem));
        }

        public List<ProjectPlanItem> ToProjectPlanList()
        {
            return projectPlanList;
        }

        public ProjectPlan CompareWithWorkReport(List<WorkReportItem> workReport)
        {
            ignoredProjectList = workReport
                .Where(workReportItem => IsNotExistInProjectPlanList(workReportItem.Name))
                .ToList()
                .Select(workReportItem => ToProjectPlanItem(workReportItem))
                .ToList();
            return this;
        }

        private ProjectPlanItem ToProjectPlanItem(WorkReportItem workReportItem)
        {
            ProjectPlanItem projectPlanItem = new ProjectPlanItem();
            projectPlanItem.ProjectName = workReportItem.Name;
            projectPlanItem.ResponsiblePersonnel = workReportItem.ResponsiblePersonnel;
            projectPlanItem.HostDivision = workReportItem.HostDivision;
            projectPlanItem.RequirementDepartment = workReportItem.BizDepartment;
            return projectPlanItem;
        }

        private bool IsNotExistInProjectPlanList(string projectName)
        {
            return !projectPlanList.Any(projectPlanItem => projectPlanItem.ProjectName.Trim().Equals(projectName));
        }

        public ProjectPlan AppendToProjectPlan()
        {
            desolator.SetCellValues(0, ignoredProjectList);
            return this;
        }
    }
}
