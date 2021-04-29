using System;
using CommandLine;
using CommandLine.Text;

namespace Minerva
{
    class Options
    {
        [Option('d', "dir", MetaValue = "FILE", Required = true, HelpText = "请输入文件夹路径")]
        public string ReportDir { get; set; }

    }
}
