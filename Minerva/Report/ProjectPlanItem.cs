using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Report
{

    //序号	项目名称	所属部门	创新类别	项目简介	立项年度	投产日期	项目类型	费用预算	数字化转型	状态	部门	负责人	激励建议
    class ProjectPlanItem
    {
        public int Sequence { get; set; }
        public string ProjectName { get; set; }
        public string RequirementDepartment { get; set; }
        public InnovationType InnovationType { get; set; }
        public string Breif { get; set; }
        public string YearOfApproval { get; set; }
        public string DateOfRelease { get; set; }
        public ProjectType ProjectType { get; set; }
        public string Budget { get; set; }
        public string IsDigitalTransformation { get; set; }
        public string Status { get; set; }
        public Division HostDivision { get; set; }
        public string ResponsiblePersonnel { get; set; }
        public string MotivationalSuggestion { get; set; }

    }
}
