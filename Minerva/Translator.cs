using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva
{
    using Report;

    class Translator
    {

        private static readonly Translator instance = new Translator();



        static Translator()
        {
        }

        private Translator()
        {
        }

        public static Translator Instance
        {
            get
            {
                return instance;
            }
        }


        private Dictionary<object, string> dict = new Dictionary<object, string>()
        {
            [Division.Division1st] = "开发一部",
            [Division.Division2nd] = "开发二部",
            [Division.Division3rd] = "开发三部",
            [Division.DivisionDataScience] = "数据分析部",


            [TaskType.RegularTechProject] = "科技项目（一般）",
            [TaskType.AgileTechProject] = "科技项目（快捷）",
            [TaskType.ImportantRequirement] = "重要需求",
            [TaskType.RegularRequirement] = "一般需求",
            [TaskType.DataMiningProject] = "数据挖掘及建模项目",
            [TaskType.TestProject] = "测试项目",
            [TaskType.Other] = "其他",


        };


        public string ToValueForHuman(object o)
        {
            if (dict[o] == null)
            {
                return o.ToString();
            }
            return dict[o];
        }

    }
}
