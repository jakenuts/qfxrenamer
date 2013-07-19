#region Using

using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using CommandLine;
using StaticVoid.Core.IO;

#endregion

namespace QFXRenamer
{
    internal class Program
    {
        // <INTU.BID>6526 
        public static Regex BankNumExpression = new Regex(@"<INTU.BID>(?<intubid>[^(\r|\</)]*)",
                                                          RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // <ACCTID>4339930014105821
        public static Regex BankNameExpression = new Regex(@"<ORG>(?<org>[^(\r|\</)]*)",
                                                           RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // <ACCTID>4339930014105821
        public static Regex AccountNumberExpression = new Regex(@"<ACCTID>(?<account>[^(\r|\</)]*)",
                                                                RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // <DTSTART>20120605120000
        public static Regex DateStartExpression = new Regex(@"<DTSTART>(?<datestart>[^(\r|\</)]*)",
                                                            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // <DTEND>20120919120000
        public static Regex DateEndExpression = new Regex(@"<DTEND>(?<dateend>[^(\r|\</)]*)",
                                                          RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static void Main(string[] args)
        {
            var options = new Options();

            if (!CommandLineParser.Default.ParseArguments(args, options))
                return;

            ReadFilesIn(options.RootDirectory);
        }


        public static void ReadFilesIn(string rootDirectory)
        {
            var rd = new DirectoryInfo(rootDirectory);

            if (!rd.Exists)
            {
                Console.Error.WriteLine("'" + rootDirectory + "' is not a valid directory.");
                return;
            }

            var webConnectFiles = rd.Files(f => f.Name.EndsWith(".qfx") || f.Name.EndsWith(".qbo"));

            foreach (var webConnectFile in webConnectFiles)
            {
                RenameIntuitFile(webConnectFile);
            }
        }

        private static void RenameIntuitFile(FileInfo qfxFile)
        {
            var text = File.ReadAllText(qfxFile.FullName);
            
            var bidMatch = BankNumExpression.Match(text);
            var orgMatch = BankNameExpression.Match(text);
            var accountMatch = AccountNumberExpression.Match(text);
            var dateStartMatch = DateStartExpression.Match(text);
            var dateEndMatch = DateEndExpression.Match(text);

            if (!accountMatch.Success || !dateStartMatch.Success || !dateEndMatch.Success)
            {
                Debug.WriteLine(text);

                Console.Error.WriteLine("'" + qfxFile.FullName + "' could not be parsed.");
                return;
            }

            var bankName = orgMatch.Groups["org"].Value.Trim();

            if (bankName == string.Empty)
                bankName = bidMatch.Groups["intubid"].Value.Trim();

            if (bankName == string.Empty)
                bankName = "UnknownBank";

            var accountName = accountMatch.Groups["account"].Value.Trim();
            var startDate = dateStartMatch.Groups["datestart"].Value.Trim();
            var endDate = dateEndMatch.Groups["dateend"].Value.Trim();

            var sd = startDate.Substring(0, 4) + "-"
                     + startDate.Substring(4, 2) + "-"
                     + startDate.Substring(6, 2);

            var ed = endDate.Substring(0, 4) + "-"
                     + endDate.Substring(4, 2) + "-"
                     + endDate.Substring(6, 2);

            var filename = string.Format(
                "{0} {1} {2} to {3}{4}",
                bankName,
                accountName,
                sd,
                ed,
                qfxFile.Extension);

            var target = Path.Combine(qfxFile.DirectoryName, filename);

            if (String.Compare(target, qfxFile.FullName, StringComparison.InvariantCultureIgnoreCase) == 0)
                return;

            var rename = target;
            var count = 0;

            while (File.Exists(rename))
            {
                rename = target.Replace(qfxFile.Extension, " (" + (++count) + ")" + qfxFile.Extension);
            }

            Debug.WriteLine("Renaming " + qfxFile.Name + " to " + rename);

            qfxFile.MoveTo(rename);
        }
    }

    public class Options
    {
        [Option("d", "directory", Required = true, HelpText = "Please provide a directory to search")]
        public string RootDirectory { get; set; }
    }
}