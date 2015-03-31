using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using CosmosTable;
using Ookii.CommandLine;

namespace CosmosTable.Command
{
    class MyArguments
    {
        [CommandLineArgument(Position = 0, IsRequired = true)]
        public string Folder { get; set; }
        //[CommandLineArgument(Position = 1)]
        //public int OptionalArgument { get; set; }
        //[CommandLineArgument]
        //public DateTime NamedArgument { get; set; }
        //[CommandLineArgument]
        //public bool SwitchArgument { get; set; }
    }

    class CommandProgram
    {
        static void RunApplication(MyArguments args)
        {
            var compiler = new Compiler();
            var folder = args.Folder;
            var lastSlash = folder.LastIndexOf("/", StringComparison.Ordinal);
            var dir = lastSlash == -1 ? "./" : folder.Substring(0, lastSlash + 1);
            var pattern = folder.Substring(lastSlash + 1, folder.Length - lastSlash - 1);
            foreach (var filePath in Directory.GetFiles(dir, pattern))
            {
                if (compiler.Compile(filePath))
                {
                    Console.Write("[Compiled Excel]: {0}", filePath);
                }
            }
        }
        static void Main(string[] args)
        {
            CommandLineParser parser = new CommandLineParser(typeof(MyArguments));
            try
            {
                MyArguments arguments = (MyArguments)parser.Parse(args);
                RunApplication(arguments);
            }
            catch (CommandLineArgumentException ex)
            {
                Console.WriteLine(ex.Message);
                parser.WriteUsageToConsole();
            }

        }
    }
}
