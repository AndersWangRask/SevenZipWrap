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
        /// <summary>
        /// Compress multiple files and directories to a 7z file.
        /// </summary>
        /// <param name="SourcePaths">The enumerable paths of files and directories to be compressed.</param>
        /// <param name="DestinationPath">The destination 7z file for the compressed files and directories.</param>
        public void CompressFiles(IEnumerable<string> SourcePaths, string DestinationPath)
        {
            CompressFiles(string.Join(" ", SourcePaths.Select(si => encaseInQuotationMarks(si))), DestinationPath);
        }

        /// <summary>
        /// Compress a file or direcotry to a 7z file.
        /// </summary>
        /// <param name="SourcePath">The path of the file or directory to be compressed.</param>
        /// <param name="DestinationPath">The destination path of the 7z file for the compressed fily or directory</param>
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

        /// <summary>
        /// Decompress the contents of the 7z file to a path.
        /// </summary>
        /// <param name="SourcePath">The path of the 7z file to decompress.</param>
        /// <param name="DestinationPath"></param>
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

            SourcePath = SourcePath.Trim();

            if (!File.Exists(SourcePath))
            {
                throw new ApplicationException($"Source 7z file was not found. File path was: {SourcePath}");
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

        /// <summary>
        /// If SevenZip executable path has not been set or has not been discovered,
        /// this will throw an exception.
        /// </summary>
        public void ThrowIfSevenZipNotInstalled()
        {
            if (!SevenZipInstalled)
            {
                throw new InvalidOperationException("7-Zip is not installed. Cannot continue.");
            }
        }

        /// <summary>
        /// Will return a Process Handle for the SevenZip executable.
        /// </summary>
        /// <param name="arguments">The (optional) arguments for the executable.</param>
        /// <returns>The Process Handle.</returns>
        protected Process getSevenZipProcess(string arguments)
        {
            //-->
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

        /// <summary>
        /// Will encase/wrap a string in quotation marks,
        /// if they are not already there.
        /// </summary>
        /// <param name="input">The string to encase.</param>
        /// <returns>A new string value encased in quotation marks.</returns>
        protected string encaseInQuotationMarks(string input)
        {
            if (input == null)
            {
                //-->
                return null;
            }

            input =
                (input.StartsWith("\"") ? "" : "\"") +
                input.Trim() +
                (input.EndsWith("\"") ? "" : "\"");

            //-->
            return input;
        }

        /// <summary>
        /// Indicates whether SevenZip is installed on this computer.
        /// </summary>
        /// <remarks>
        /// This depends on whether a SevenZip executable path has been found in the registry,
        /// or that a valid SevenZip executable path has been set.
        /// <seealso cref="SevenZipApplicationPath"/>
        /// </remarks>
        public bool SevenZipInstalled => SevenZipApplicationPath != null;

        /// <summary>
        /// Get or Set the path to the SevenZip executable.
        /// If not Set, the Get will look up the path in the registry.
        /// </summary>
        /// <remarks>
        /// If the value is set, there is a limited test of validity. The file at path has to exist.
        /// If the value is not set, it is looked up in the registry. If SevenZip is not installed, there should be no registry entry.
        /// <seealso cref="getSevenZipApplicationPath"/>
        /// </remarks>
        public string SevenZipApplicationPath
        {
            get
            {
                if (!_sevenZipApplicationPathSet)
                {
                    _sevenZipApplicationPath = getSevenZipApplicationPath();
                    _sevenZipApplicationPathSet = true;
                }

                //-->
                return _sevenZipApplicationPath;
            }
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    _sevenZipApplicationPath = null;
                    _sevenZipApplicationPathSet = false;
                }
                else
                {
                    string newPath = value.Trim();

                    if (!File.Exists(newPath))
                    {
                        throw new ApplicationException($"SevenZip File Executable path is not valid. The file was not faound. Path was: {newPath}");
                    }

                    _sevenZipApplicationPath = newPath;
                    _sevenZipApplicationPathSet = true;
                }
            }
        }
        protected string _sevenZipApplicationPath = null;
        protected bool _sevenZipApplicationPathSet = false;

        /// <summary>
        /// Looks up the SevenZip executable from the registry.
        /// Requires SevenZip to be installed.
        /// Otherwise will return null.
        /// </summary>
        /// <returns>
        /// The full path to the SevenZip executable if found.
        /// Otherwise null.
        /// </returns>
        protected string getSevenZipApplicationPath()
        {
            const string regPath = @"SOFTWARE\7-Zip\";
            const string valueName = "Path";

            RegistryView registryView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;

            string path =
                (string)(
                    RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView).OpenSubKey(regPath)?.GetValue(valueName)
                    ??
                    RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, registryView).OpenSubKey(regPath)?.GetValue(valueName));

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