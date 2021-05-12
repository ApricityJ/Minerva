namespace Minerva.Weekly
{
    internal class WeeklyItem
    {
        //序号	项目/需求/管理名称	任务类型	当前进展	主办部门	协办部门	业务部门	负责人员	参与人员	进度计划
        public int Sequence { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string CurrentProgress { get; set; }
        public string HostDivision { get; set; }
        public string CoHostDivision { get; set; }
        public string BizDepartment { get; set; }
        public string ResponsiblePersonnel { get; set; }
        public string Participants { get; set; }
        public string Schedule { get; set; }

    }
}
