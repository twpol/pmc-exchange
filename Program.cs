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

            GetUnreadMessages(service).ForEachAsync(message => WriteEmailLine(message)).Wait();
            GetFlaggedMessages(service).ForEachAsync(message => WriteEmailLine(message)).Wait();

            return 0;
        }

        static HashSet<string> EmailsSeen = new HashSet<string>();

        static void WriteEmailLine(EmailMessage message)
        {
            if (EmailsSeen.Contains(message.Id.UniqueId))
                return;
            Console.WriteLine(EmailToJson(message).ToString());
            EmailsSeen.Add(message.Id.UniqueId);
        }

        static JsonObject EmailToJson(EmailMessage message)
        {
            return new JsonObject(
                new KeyValuePair<string, JsonValue>("source", "pmc-exchange"),
                new KeyValuePair<string, JsonValue>("type", "email"),
                new KeyValuePair<string, JsonValue>("id", message.Id.UniqueId),
                new KeyValuePair<string, JsonValue>("datetime", message.DateTimeReceived.ToString("O")),
                new KeyValuePair<string, JsonValue>("subject", message.Subject),
                new KeyValuePair<string, JsonValue>("flagged", message.Flag.FlagStatus != ItemFlagStatus.NotFlagged),
                new KeyValuePair<string, JsonValue>("completed", message.Flag.FlagStatus == ItemFlagStatus.Complete),
                new KeyValuePair<string, JsonValue>("read", message.IsRead)
            );
        }

        static IObservable<EmailMessage> GetUnreadMessages(ExchangeService service)
        {
            return Observable.Create<EmailMessage>(
                async observer =>
                {
                    // Find "Inbox" and all child folders.
                    var folders = new List<Folder>();
                    folders.Add(await Folder.Bind(service, WellKnownFolderName.Inbox));
                    var folderView = new FolderView(10)
                    {
                        Traversal = FolderTraversal.Deep,
                    };
                    {
                        FindFoldersResults folder;
                        do
                        {
                            folder = await folders[0].FindFolders(folderView);
                            folders.AddRange(folder);
                            folderView.Offset = folder.NextPageOffset ?? 0;
                        } while (folder.MoreAvailable);
                    }

                    // Find all items that are unread.
                    var unreadFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And)
                    {
                        new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false),
                        new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note"),
                    };
                    var unreadView = new ItemView(1000)
                    {
                        PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.DateTimeReceived, ItemSchema.Subject, ItemSchema.Flag, EmailMessageSchema.IsRead),
                    };

                    foreach (var folder in folders)
                    {
                        FindItemsResults<Item> unread;
                        do
                        {
                            unread = await folder.FindItems(unreadFilter, unreadView);
                            foreach (var item in unread.Items)
                            {
                                observer.OnNext(item as EmailMessage);
                            }
                            unreadView.Offset = unread.NextPageOffset ?? 0;
                        } while (unread.MoreAvailable);
                    }

                    observer.OnCompleted();
                }
            );
        }

        static IObservable<EmailMessage> GetFlaggedMessages(ExchangeService service)
        {
            return Observable.Create<EmailMessage>(
                async observer =>
                {
                    // ItemSchema.Flag does not seem to be searchable
                    var PidTagFlagStatus = new ExtendedPropertyDefinition(0x1090, MapiPropertyType.Integer);

                    // Find Outlook's own search folder "AllItems", which includes all folders in the account.
                    var allItemsView = new FolderView(10);
                    var allItems = await service.FindFolders(WellKnownFolderName.Root,
                        new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "AllItems"),
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
                        PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.DateTimeReceived, ItemSchema.Subject, ItemSchema.Flag, EmailMessageSchema.IsRead),
                    };

                    FindItemsResults<Item> flagged;
                    do
                    {
                        flagged = await allItems.Folders[0].FindItems(flaggedFilter, flaggedView);
                        foreach (var item in flagged.Items)
                        {
                            observer.OnNext(item as EmailMessage);
                        }
                        flaggedView.Offset = flagged.NextPageOffset ?? 0;
                    } while (flagged.MoreAvailable);

                    observer.OnCompleted();
                }
            );
        }
    }
}
