using System;
using System.Collections.Generic;
using System.IO;
using System.Json;
using System.Reactive.Linq;
using CommandLine;
using Microsoft.Exchange.WebServices.Data;
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

        static int Run(IConfigurationRoot config, Options options)
        {
            var service = new ExchangeService(ExchangeVersion.Exchange2016);
            service.Credentials = new WebCredentials(config["username"], config["password"]);
            service.AutodiscoverUrl(config["email"], redirectionUri => new Uri(redirectionUri).Scheme == "https");

            GetFlaggedMessages(service).ForEachAsync(message =>
            {
                var item = new JsonObject(
                    new KeyValuePair<string, JsonValue>("source", "pmc-exchange"),
                    new KeyValuePair<string, JsonValue>("type", "email"),
                    new KeyValuePair<string, JsonValue>("completed", message.Flag.FlagStatus == ItemFlagStatus.Complete),
                    new KeyValuePair<string, JsonValue>("rank", 0),
                    new KeyValuePair<string, JsonValue>("datetime", message.DateTimeReceived.ToString("O")),
                    new KeyValuePair<string, JsonValue>("subject", message.Subject)
                );
                Console.WriteLine(item.ToString());
            }).Wait();

            return 0;
        }

        static IObservable<Item> GetFlaggedMessages(ExchangeService service)
        {
            return Observable.Create<Item>(
                async observer =>
                {
                    var PidTagFolderType = new ExtendedPropertyDefinition(0x3601, MapiPropertyType.Integer);
                    var PidTagFlagStatus = new ExtendedPropertyDefinition(0x1090, MapiPropertyType.Integer);

                    // Find Outlook's own search folder "AllItems", which includes all folders in the account.
                    var allItemsView = new FolderView(10);
                    var allItems = await service.FindFolders(WellKnownFolderName.Root,
                        new SearchFilter.SearchFilterCollection(LogicalOperator.And)
                        {
                            new SearchFilter.IsEqualTo(PidTagFolderType, "2"),
                            new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "AllItems"),
                        },
                        allItemsView);

                    if (allItems.Folders.Count != 1)
                    {
                        throw new MissingMemberException("AllItems");
                    }

                    // Find the Junk folder.
                    var junkFolder = await Folder.Bind(service, WellKnownFolderName.JunkEmail);

                    // Find all items that are flagged and not in the Junk folder.
                    var flaggedFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And)
                    {
                        new SearchFilter.Exists(PidTagFlagStatus),
                        new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note"),
                        new SearchFilter.IsNotEqualTo(ItemSchema.ParentFolderId, junkFolder.Id.UniqueId),
                    };
                    var flaggedView = new ItemView(1000)
                    {
                        PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.DateTimeReceived, ItemSchema.Flag, ItemSchema.Subject),
                    };

                    FindItemsResults<Item> flagged;
                    do
                    {
                        flagged = await allItems.Folders[0].FindItems(flaggedFilter, flaggedView);
                        foreach (var item in flagged.Items)
                        {
                            observer.OnNext(item);
                        }
                        flaggedView.Offset = flagged.NextPageOffset ?? 0;
                    } while (flagged.MoreAvailable);

                    observer.OnCompleted();
                }
            );
        }
    }
}
