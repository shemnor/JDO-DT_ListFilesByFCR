using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace JDO_DT_ListFilesByFCR
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            char delimeter = ","[0];
            string[] requestedFCRs = textBox1.Text.Split(delimeter);

            Dictionary<string, List<string>> fcrFileNames = new Dictionary<string, List<string>>();
            List<FCR> FCRs = new List<FCR>();


            //Dictionary<string, List<string>> fileNamesA = new Dictionary<string, List<string>>();
            getDwgAndMdlFilesByFcrs(@"C:\Users\SZCZ1360\OneDrive Corp\Atkins Ltd\UK1226_SDLT - Documents\04_HD\01 Unit 1\01 HDAB\01 FCR", requestedFCRs, ref FCRs);

            //Dictionary<string, List<string>> fileNamesB = new Dictionary<string, List<string>>();
            getDwgAndMdlFilesByFcrs(@"C:\Users\SZCZ1360\OneDrive Corp\Atkins Ltd\UK1226_SDLT - Documents\04_HD\01 Unit 1\02 HDCD\01 FCR", requestedFCRs, ref FCRs);

            //get latest revisions only
            Dictionary<string, file> latestRevisions = returnLatestFiles(FCRs);

            //remove duplicates and old versions
            removeOldRevisionsAndDuplicates(latestRevisions, ref FCRs);

            //Start Excel and get Application object.
            Excel.Application oXL = new Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            Excel._Workbook oWB = (Excel._Workbook)(oXL.Workbooks.Add());
            Excel._Worksheet oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            oSheet.Cells[1, 1] = "FCR";
            oSheet.Cells[1, 2] = "fileName";

            printDictInExcel(oSheet, FCRs);
        }

        private static void getDwgAndMdlFilesByFcrs(string searchDir, string[] fcrNumbers, ref List<FCR> fcrFileNames)
        {
            List<string> dirs = FolderTools.Class1.ListDirs(searchDir);
            //Dictionary<string, List<string>> fcrFileNames = new Dictionary<string, List<string>>();
            foreach (string folder in dirs)
            {
                for (int i = 0; i < fcrNumbers.Length; i++)
                {
                    string folderName = folder.Substring(folder.LastIndexOf("\\") + 1);
                    if (folderName.Length > 6)
                    {
                        string FCR = folderName.Substring(0, 6);
                        if (fcrNumbers[i].Equals(FCR))
                        {
                            //get info
                            string dwgFolder = folder + @"\DRAWINGS";
                            string mdlFolder = folder + @"\IFC";
                            List<string> fileNames = new List<string>(getFileNames(dwgFolder));
                            fileNames.AddRange(getFileNames(mdlFolder, "*.ifc"));

                            List<file> newFiles = new List<file>();
                            foreach (string rawFileName in fileNames)
                            {
                                file newFile = new file();
                                newFile.sketchPrefix = rawFileName.Substring(0, 9);
                                newFile.fileName = rawFileName.Substring(10, rawFileName.LastIndexOf("-") - 10);
                                char letter = char.Parse(rawFileName.Substring(rawFileName.LastIndexOf("-") + 1, 1));
                                newFile.revisionLetterIndex = char.ToUpper(letter) - 64;
                                newFile.revisionNumber = int.Parse(rawFileName.Substring(rawFileName.LastIndexOf(".") - 2, 2));

                                newFiles.Add(newFile);
                            }

                            //save info
                            FCR newFCR = new FCR();
                            newFCR.FcrNumber = FCR;
                            newFCR.Files = newFiles;
                            fcrFileNames.Add(newFCR);
                        }
                    }
                }
            }
        }

        private static Dictionary<string, file> returnLatestFiles(List<FCR> FCRs)
        {
            Dictionary<string, file> latestRevisions = new Dictionary<string, file>();

            foreach (FCR FCR in FCRs)
            {
                foreach (file file in FCR.Files)
                {
                    if (latestRevisions.ContainsKey(file.fileName))
                    {
                        //get infor for comparison
                        file latestFile = latestRevisions[file.fileName];
                        int latestLetterIndex = latestFile.revisionLetterIndex;
                        int latestNumber = latestFile.revisionNumber;

                        //if letter index is lower then the latest then replace the full revision
                        if (latestLetterIndex < file.revisionLetterIndex)
                        {
                            latestFile.revisionLetterIndex = file.revisionLetterIndex;
                            latestFile.revisionNumber = file.revisionNumber;
                        }

                        //if the same letter index but lower number then replace the full revision
                        else if (latestLetterIndex.Equals(file.revisionLetterIndex))
                        {
                            if (latestNumber < file.revisionNumber)
                            {
                                latestFile.revisionLetterIndex = file.revisionLetterIndex;
                                latestFile.revisionNumber = file.revisionNumber;
                            }
                        }
                    }
                    else
                    {
                        latestRevisions.Add(file.fileName, new file(file));
                    }
                }
            }
            return latestRevisions;
        }

        private static void removeOldRevisionsAndDuplicates(Dictionary<string, file> latestRevisions, ref List<FCR> FCRs)
        {
            foreach (string latestFileName in latestRevisions.Keys)
            {
                bool duplicate = false;
                file latestFile = latestRevisions[latestFileName];

                foreach (FCR FCR in FCRs)
                {
                    List<int> toDelete = new List<int>();

                    //find files to delete
                    for (int i = 0; i < FCR.Files.Count; i++)
                    {
                        file file = FCR.Files[i];
                        if (file.fileName.Equals(latestFileName))
                        {
                            //if the file is the latest file and is not a duplicate, mark as found
                            if (file.revisionLetterIndex.Equals(latestFile.revisionLetterIndex) && file.revisionNumber.Equals(latestFile.revisionNumber) && !duplicate)
                            {
                                duplicate = true;
                            }
                            //if its not most up to date or its a duplicate then remove
                            else
                            {
                                toDelete.Add(i);
                            }
                        }
                    }

                    //delete outdated files
                    for (int i = toDelete.Count - 1; i >= 0; i--)
                    {
                        FCR.Files.RemoveAt(toDelete[i]);
                    }
                }
            }
        }

        private static string[] getFileNames(string dirPath, string wildcard = "*")
        {
            string[] dirFiles = Directory.GetFiles(dirPath, wildcard);

            for (int k = 0; k < dirFiles.Length; k++)
            {
                dirFiles[k] = dirFiles[k].Substring(dirFiles[k].LastIndexOf("\\") + 1);
            }
            return dirFiles;
        }

        private static void printDictInExcel(Excel._Worksheet oSheet, List<FCR> FCRs)
        {
            Excel.Range last = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int row = last.Row + 1;

            foreach (FCR FCR in FCRs)
            {
                oSheet.Cells[row, 1] = FCR.FcrNumber;
                foreach (file file in FCR.Files)
                {
                    string letter = ((char)(file.revisionLetterIndex + 64)).ToString();
                    if (file.revisionNumber < 10)
                    {
                        oSheet.Cells[row, 2] = string.Format("{0} {1}-{2}.0{3}", file.sketchPrefix, file.fileName, letter, file.revisionNumber);
                    }
                    else
                    {
                        oSheet.Cells[row, 2] = string.Format("{0} {1}-{2}.{3}", file.sketchPrefix, file.fileName, letter, file.revisionNumber);
                    }
                    row += 1;
                }
            }
        }
    }
    public class file
    {
        public file()
        {

        }
        public file(file previousFile)
        {
            sketchPrefix = previousFile.sketchPrefix;
            fileName = previousFile.fileName;
            revisionLetterIndex = previousFile.revisionLetterIndex;
            revisionNumber = previousFile.revisionNumber;
        }
        public string sketchPrefix { get; set; }
        public string fileName { get; set; }
        public int revisionLetterIndex { get; set; }
        public int revisionNumber { get; set; }

    }

    public class FCR
    {
        public string FcrNumber { get; set; }
        public List<file> Files { get; set; }

    }
}
