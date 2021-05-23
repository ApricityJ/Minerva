using System;
using CommandLine;
using CommandLine.Text;

namespace Minerva
{
    internal class Options
    {
        [Option('d', "dir", MetaValue = "FILE", Required = true, HelpText = "请输入文件夹路径")]
        public string ReportDir { get; set; }

        [Option('p', "phase", MetaValue = "INT", Required = true, HelpText = "请输入当前阶段")]
        public int Phase { get; set; }
    }
}
