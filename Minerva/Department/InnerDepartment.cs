namespace Minerva.Department
{
    /// <summary>
    /// 科技部内部部门，实际应为10个
    /// 排序为开发1，2，3，4（数据），5（测试）
    /// 系统/网络
    /// 产品/综合 
    /// 按5，2，2排列
    /// （运行部：我不配拥有姓名？）
    /// （综合部：周报是什么？）
    /// 所以只剩下8个交了周报的部门
    /// 
    /// 这是一种偷懒的做法，不偷懒应该：
    /// 1、定义一个InnerDept的实体类，属性包含id，名称，是否开发部门，etc
    /// 2、定义一个叫IT&PMDept的类，用来管理这些部门
    /// 一想就觉得不乐意，下次一定
    /// </summary>
    public enum InnerDepartment
    {
        DevDivision1st,   //开发一部
        DevDivision2nd,   //开发一部 
        DevDivision3rd,   //开发三部
        DevDataSci,       //数据分析
        Test,             //测试
        System,           //系统
        Network,          //网络
        Product,          //产品
    }
}
