using System;
using System.IO;
using CommandLine;
using Microsoft.Extensions.Configuration;

namespace pmc.exchange
{
    class Program
    {
        class Options
        {
            [Option(HelpText = "Display details of what the program is doing.")]
            public bool Verbose { get; set; } = false;

            [Option(HelpText = "Specify the configuration file to use (default: config.json).")]
            public FileInfo Config { get; set; } = new FileInfo("config.json");
        }

        static int Main(string[] args)
        {
            return Parser.Default.ParseArguments<Options>(args)
                .MapResult(
                    options => Run(new ConfigurationBuilder()
                        .AddJsonFile(options.Config.FullName, true)
                        .Build(), options),
                    _ => 1
                );
        }

        static int Run(IConfigurationRoot configurationRoot, Options options)
        {
            return 0;
        }
    }
}
