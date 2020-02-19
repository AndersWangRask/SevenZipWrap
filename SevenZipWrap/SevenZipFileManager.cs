using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace SevenZipWrap
{
    public class SevenZipFileManager
    {
        public void CompressFiles(IEnumerable<string> SourcePaths, string DestinationPath)
        {
            CompressFiles(string.Join(" ", SourcePaths.Select(si => encaseInQuotationMarks(si))), DestinationPath);
        }

        public void CompressFiles(string SourcePath, string DestinationPath)
        {
            ThrowIfSevenZipNotInstalled();

            if (string.IsNullOrWhiteSpace(SourcePath))
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(DestinationPath))
            {
                throw new ArgumentNullException(nameof(DestinationPath));
            }

            string arguments = " a " + encaseInQuotationMarks(DestinationPath) + " " + encaseInQuotationMarks(SourcePath) + " -mx=7 -mmt=off -mtc=on";
            Process process = getSevenZipProcess(arguments);
            process.Start();

            while (!process.HasExited)
            {
                Thread.Sleep(100);
            }

            if (process.ExitCode != 0)
            {
                string msg =
                    "7-Zip Compress operation did not complete succesfully. Exit Code: " + process.ExitCode +
                    "\nApplication Path: " + SevenZipApplicationPath +
                    "\nArguments: " + arguments +
                    "\nOutput:\n" + process.StandardOutput?.ReadToEnd();

                Debug.WriteLine("ERROR:\n" + msg + "\n");

                throw new ApplicationException(msg);
            }
            else
            {
                Debug.WriteLine("Completed Compress Operation\nOutput:\n" + process.StandardOutput?.ReadToEnd() + "\n");
            }
        }

        public void DecompressFiles(string SourcePath, string DestinationPath)
        {
            ThrowIfSevenZipNotInstalled();

            if (string.IsNullOrWhiteSpace(SourcePath))
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(DestinationPath))
            {
                throw new ArgumentNullException(nameof(DestinationPath));
            }

            string arguments = " x " + encaseInQuotationMarks(SourcePath) + " -o" + encaseInQuotationMarks(DestinationPath) + " -y";
            Process process = getSevenZipProcess(arguments);
            process.Start();

            while (!process.HasExited)
            {
                Thread.Sleep(100);
            }

            if (process.ExitCode != 0)
            {
                string msg =
                    "7-Zip Decompress operation did not complete succesfully. Exit Code: " + process.ExitCode +
                    "\nApplication Path: " + SevenZipApplicationPath +
                    "\nArguments: " + arguments +
                    "\nOutput:\n" + process.StandardOutput?.ReadToEnd();

                Debug.WriteLine("ERROR:\n" + msg + "\n");

                throw new ApplicationException(msg);
            }
            else
            {
                Debug.WriteLine("Completed Decompress Operation\nOutput:\n" + process.StandardOutput?.ReadToEnd() + "\n");
            }
        }

        public void ThrowIfSevenZipNotInstalled()
        {
            if (!SevenZipInstalled)
            {
                throw new InvalidOperationException("7-Zip is not installed. Cannot continue.");
            }
        }

        protected Process getSevenZipProcess(string arguments)
        {
            return
                new Process
                {
                    StartInfo =
                    {
                        Arguments = arguments,
                        FileName = SevenZipApplicationPath,
                        UseShellExecute = false,
                        RedirectStandardOutput = true
                    }
                };
        }

        protected string encaseInQuotationMarks(string input)
        {
            input =
                (input.StartsWith("\"") ? "" : "\"") +
                input.Trim() +
                (input.EndsWith("\"") ? "" : "\"");

            return input;
        }

        public bool SevenZipInstalled => SevenZipApplicationPath != null;

        public string SevenZipApplicationPath
        {
            get
            {
                if (!_sevenZipApplicationPathSet)
                {
                    _sevenZipApplicationPath = getSevenZipApplicationPath();
                    _sevenZipApplicationPathSet = true;
                }

                return _sevenZipApplicationPath;
            }
        }
        protected string _sevenZipApplicationPath = null;
        protected bool _sevenZipApplicationPathSet = false;

        protected string getSevenZipApplicationPath()
        {
            const string regPath = @"SOFTWARE\7-Zip\";
            const string valueName = "Path";

            RegistryView registryView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;

            string path =
                (string)
                (RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView).OpenSubKey(regPath).GetValue(valueName)
                ??
                RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, registryView).OpenSubKey(regPath).GetValue(valueName));

            if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
            {
                path = Path.Combine(path, "7z.exe");

                if (File.Exists(path))
                {
                    //-->
                    return path;
                }
            }

            //-->
            return null;
        }
    }
}
