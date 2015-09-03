using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using CosmosTable;
using Neo.IronLua;
using Ookii.CommandLine;

namespace CosmosTable.Command
{
    class MyArguments
    {
        [CommandLineArgument()]
        public string Folder { get; set; }

        [CommandLineArgument(Position = 0)]
        public string Script { get; set; }
        //[CommandLineArgument(Position = 1)]
        //public int OptionalArgument { get; set; }
        //[CommandLineArgument]
        //public DateTime NamedArgument { get; set; }
        //[CommandLineArgument]
        //public bool SwitchArgument { get; set; }
    }

    public class CommandProgram
    {
        static void RunApplication(MyArguments args)
        {
            var compiler = new Compiler();
            var folder = args.Folder;
            if (!string.IsNullOrEmpty(folder))
            {
                Console.WriteLine("Execute Folder: {0}", folder);
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

            if (!string.IsNullOrEmpty(args.Script))
            {
                if (!File.Exists(args.Script))
                {
                    Console.WriteLine("[Error] Not found script:　{0}", args.Script);
                }
                else
                {
                    Console.WriteLine("Execute Script: {0}", args.Script);
                    using (Lua l = new Lua()) // create the lua script engine
                    {
                        dynamic g = l.CreateEnvironment(); // create a environment
                        g.dochunk(File.ReadAllText(args.Script)); // create a variable in lua
                        Console.WriteLine(g.abc); // access a variable in c#
                                                  //g.dochunk("function add(b) return b + 3; end;", "test.lua"); // create a function in lua
                                                  //Console.WriteLine("Add(3) = {0}", g.add(3)); // call the function in c#
                    }
                }

            }
        }

        public static void Run(string[] args)
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
        static void Main(string[] args)
        {
            Run(args);
        }
    }
}
