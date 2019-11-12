using System;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using ImapFolderSubscriptionGuard.Properties;
using IniParser;
using IniParser.Model;
using IniParser.Model.Configuration;
using IniParser.Parser;
using ThrottleDebounce;

#nullable enable

namespace ImapFolderSubscriptionGuard {

    internal class SubscriptionGuard: IDisposable {

        private const string SUBSCRIPTION_CONFIGURATION_FILENAME = "HIWATER.MRK";
        private static readonly ManualResetEvent KEEP_PROGRAM_RUNNING = new ManualResetEvent(false);

        private readonly string mailboxDirectory;
        private readonly StringCollection foldersToUnsubscribeFrom;
        private readonly FileSystemWatcher fileSystemWatcher;

        private readonly FileIniDataParser initializationDataMapper = new FileIniDataParser(new IniDataParser(
            new IniParserConfiguration {
                AssigmentSpacer = string.Empty
            }));

        private string subscriptionConfigurationPath => Path.Combine(mailboxDirectory, SUBSCRIPTION_CONFIGURATION_FILENAME);

        private static void Main(string[] args) {
            if (args.Contains("--console")) {
                using var subscriptionGuard = new SubscriptionGuard();
                subscriptionGuard.startGuarding();
                KEEP_PROGRAM_RUNNING.WaitOne();
            } else {
                ServiceBase.Run(new Service());
            }
        }

        internal SubscriptionGuard() {
            mailboxDirectory = Settings.Default.UserMailboxDirectory;
            foldersToUnsubscribeFrom = Settings.Default.UnsubscribeFromFolders;

            if (string.IsNullOrWhiteSpace(mailboxDirectory) || foldersToUnsubscribeFrom?.Count == 0) {
                Console.WriteLine($"Configuration missing, ensure {Process.GetCurrentProcess().ProcessName}.config exists in the " +
                                  "same directory as this executable.");
                Environment.Exit(1);
            }

            fileSystemWatcher = new FileSystemWatcher(mailboxDirectory, SUBSCRIPTION_CONFIGURATION_FILENAME) {
                NotifyFilter = NotifyFilters.LastWrite
            };
        }

        internal void startGuarding() {
            fixSubscriptions();

            DebouncedAction onSubscriptionFileChanged = Debouncer.Debounce(() => {
                Console.WriteLine($"\nChange detected in {subscriptionConfigurationPath}");
                fileSystemWatcher.EnableRaisingEvents = false;
                fixSubscriptions();
                fileSystemWatcher.EnableRaisingEvents = true;
            }, TimeSpan.FromMilliseconds(50));

            fileSystemWatcher.Changed += delegate { onSubscriptionFileChanged.Run(); };

            fileSystemWatcher.EnableRaisingEvents = true;

            Console.WriteLine($"Waiting for subscription changes in {subscriptionConfigurationPath}");
        }

        internal void stopGuarding() {
            fileSystemWatcher.EnableRaisingEvents = false;
        }

        public void Dispose() {
            fileSystemWatcher?.Dispose();
        }

        private void fixSubscriptions() {
            unsubscribeFromFolders();
            deleteFolders();
        }

        private void unsubscribeFromFolders() {
            lock (SUBSCRIPTION_CONFIGURATION_FILENAME) {
                IniData subscriptionConfiguration = initializationDataMapper.ReadFile(subscriptionConfigurationPath, Encoding.ASCII);
                KeyDataCollection imapSubscribedSection = subscriptionConfiguration["IMAPSubscribed"];

                bool wasConfigurationChanged = false;

                foreach (string folderToUnsubscribeFrom in foldersToUnsubscribeFrom) {
                    bool didKeyExist = imapSubscribedSection.RemoveKey(folderToUnsubscribeFrom);
                    if (didKeyExist) {
                        wasConfigurationChanged = true;
                        Console.WriteLine($"Unsubscribing from folder {folderToUnsubscribeFrom}");
                    }
                }

                if (wasConfigurationChanged) {
                    initializationDataMapper.WriteFile(subscriptionConfigurationPath, subscriptionConfiguration, Encoding.ASCII);
                    Console.WriteLine($"Saved {subscriptionConfigurationPath}");
                }
            }
        }

        private void deleteFolders() {
            foreach (string folderToUnsubscribeFrom in foldersToUnsubscribeFrom) {
                string folderPath = Path.Combine(mailboxDirectory, folderToUnsubscribeFrom + ".IMAP");
                try {
                    Directory.Delete(folderPath, true);
                    Console.WriteLine($"Deleted directory {folderPath}");
                } catch (DirectoryNotFoundException) {
                    // Good, the folder already didn't exist. Keep going.
                }
            }
        }

    }

}