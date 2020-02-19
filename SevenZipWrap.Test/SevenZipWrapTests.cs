using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SevenZipWrap;

namespace SevenZipWrap.Test
{
    [TestClass]
    public class SevenZipWrapTests
    {
        private string testDirPath = null;
        private const string sourceSubPath = "source";
        private const string destSubPath = "dest";
        private const string unpackSubPath = "unpack";
        private List<string> sourceFiles = new List<string>();

        [TestInitialize]
        public void Init()
        {
            Debug.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Init: Begin");

            //Create directory for test
            testDirPath = $@"c:\test7zw-{Guid.NewGuid()}";

            try
            {
                Directory.CreateDirectory(testDirPath);
                Debug.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Init: Created test directory: {testDirPath} ");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Init: Cannot continue. Could not create test directory. (Could be related to permissions.) Directory Path: {testDirPath}\n{ex} ");
                testDirPath = null;
                throw;
            }

            //Create sub-directories
            Directory.CreateDirectory(Path.Combine(testDirPath, sourceSubPath));
            Directory.CreateDirectory(Path.Combine(testDirPath, destSubPath));
            Directory.CreateDirectory(Path.Combine(testDirPath, unpackSubPath));

            //Create sample source-files
            foreach (int i1 in Enumerable.Range(0, 2))
            {
                string sourceFilePath = Path.Combine(testDirPath, sourceSubPath, $"source-{Guid.NewGuid()}.txt");

                using (StreamWriter fileStream = File.CreateText(sourceFilePath))
                {
                    foreach (int i2 in Enumerable.Range(0, 1000))
                    {
                        fileStream.WriteLine(Guid.NewGuid().ToString());
                    }
                }

                sourceFiles.Add(sourceFilePath);

                Debug.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Init: Created sample source file: {sourceFilePath} ");
            }


            Debug.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Init: Complete");
        }

        [TestCleanup]
        public void Clean()
        {
            Debug.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Clean: Begin");

            if (!string.IsNullOrWhiteSpace(testDirPath) && Directory.Exists(testDirPath))
            {
                Directory.Delete(testDirPath, true);
                Debug.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Clean: Complete");
            }

            Debug.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Clean: Complete");
        }

        [TestMethod]
        public void FileManager_CompressUncompressSingleFile()
        {
            //Arrange
            SevenZipFileManager sevenZip = new SevenZipFileManager();

            string source1 = sourceFiles.FirstOrDefault();
            string dest1 = Path.Combine(testDirPath, destSubPath, Path.GetFileName(source1) + ".7z");
            string unpack1 = Path.Combine(testDirPath, unpackSubPath, "1");

            Directory.CreateDirectory(unpack1);

            //Act
            sevenZip.CompressFiles(source1, dest1);
            sevenZip.DecompressFiles(dest1, unpack1);

            //Assert
            Assert.IsTrue(File.Exists(dest1));
            Assert.IsTrue(Directory.Exists(unpack1));
            Assert.IsTrue(File.Exists(Directory.GetFiles(unpack1).FirstOrDefault()));
        }

        [TestMethod]
        public void FileManager_CompressUncompressMultiFiles()
        {
            //Arrange
            SevenZipFileManager sevenZip = new SevenZipFileManager();

            IEnumerable<string> source2 = sourceFiles;
            string dest2 = Path.Combine(testDirPath, destSubPath, "2.7z");
            string unpack2 = Path.Combine(testDirPath, unpackSubPath, "2");

            Directory.CreateDirectory(unpack2);

            //Act
            sevenZip.CompressFiles(source2, dest2);
            sevenZip.DecompressFiles(dest2, unpack2);

            //Assert
            Assert.IsTrue(File.Exists(dest2));
            Assert.IsTrue(Directory.Exists(unpack2));
            Assert.IsTrue(Directory.GetFiles(unpack2).Count() == source2.Count());
        }

        [TestMethod]
        public void FileManager_CompressUncompressDirectory()
        {
            //Arrange
            SevenZipFileManager sevenZip = new SevenZipFileManager();

            string source3 = Path.Combine(testDirPath, sourceSubPath);
            string dest3 = Path.Combine(testDirPath, destSubPath, "3.7z");
            string unpack3 = Path.Combine(testDirPath, unpackSubPath, "3");

            Directory.CreateDirectory(unpack3);

            //Act
            sevenZip.CompressFiles(source3, dest3);
            sevenZip.DecompressFiles(dest3, unpack3);

            //Assert
            Assert.IsTrue(File.Exists(dest3));
            Assert.IsTrue(Directory.Exists(unpack3));
            Assert.IsTrue(Directory.Exists(Directory.GetDirectories(unpack3).FirstOrDefault()));
        }
    }
}
