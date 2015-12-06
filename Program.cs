using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.ServiceProcess;
using System.IO;
using System.Threading;

namespace UpgradeUtil
{
    class Program
    {
        static List<string> logLines = new List<string>();
        static void Main(string[] args)
        {
            int currentBuild, newestBuild;
            string release = @"6.0MR2";
            string destPath = @"C:\UpgradeUtil\Installer\CGSrv.exe";
            string ahkScript = @"C:\UpgradeUtil\UpgradeUtilAHK.exe";
            string newestBuildPath;
            string mapServerPath = @"\\buffalo2\qabuilds";
            string mapArgs = @"/user:headquarters\buildadmin build1it /persistent:yes";
            string tempString;
            bool upgradeBool;
            string moduleDefPathOld;
            string moduleDefPathNew;
            bool isModuleDefInUse = false;
            
            //Map to the buffalo server
            MapNetworkDrive(mapServerPath, mapArgs);
            
            //Close the CG client or any open browsers
            CloseProcesses();

            CreateInstallerFolder();
            currentBuild = GetLastBuildUpgrade();
            newestBuild = GetNewestBuildNumber(release);
            tempString = string.Format("Current Build: {0} Newest Build: {1}", currentBuild, newestBuild);
            logLines.Add(tempString);
            Console.WriteLine(tempString);
            
            //Check if 
            if (currentBuild == 0)
            {
                upgradeBool = false;
                logLines.Add("ChangeGear install was not found");
                Console.WriteLine("ChangeGear install was not found");
            }
            else
            {
                upgradeBool = DetermineIfUpgradeNecessary(currentBuild, newestBuild);
            }

            if (upgradeBool)
            {
                //Module def renaming to bak if it exists
                moduleDefPathOld = GetChangeGearpath() + @"Server\Views\ModuleDef.xml";
                moduleDefPathNew = GetChangeGearpath() + @"Server\Views\ModuleDef.bak";
                if (File.Exists(moduleDefPathOld))
                {
                    RenameFile(moduleDefPathOld, moduleDefPathNew);
                    isModuleDefInUse = true;
                }
                newestBuildPath = GetNewestBuildPath(release, newestBuild) + @"\DownloadImages\CGSrv.exe";
                logLines.Add("Copying installer...");
                Console.WriteLine("Copying installer...");
                try
                {
                    File.Copy(newestBuildPath, destPath, true);
                    logLines.Add("Installer successfully copied!");
                    Console.WriteLine("Installer successfully copied!");
                }
                catch(Exception)
                {
                    logLines.Add("File copy failed!");
                    Console.WriteLine("File copy failed!");
                    return;
                }

                //Wait 5 seconds after copying
                Thread.Sleep(5000);

                //Start the AutoHotKey script
                Process installer = new Process();
                installer.StartInfo.FileName = ahkScript;
                //installer.StartInfo.Verb = "runas";
                installer.Start();
                //System.Diagnostics.Process installer = System.Diagnostics.Process.Start(ahkScript);
                installer.WaitForExit();

                currentBuild = newestBuild;
                SetLastBuildUpgrade(currentBuild);

                //Wait and close the browser and the CGClient. Reset the IIS server
                Thread.Sleep(20000);
                CloseProcesses();

                Thread.Sleep(4000);
                //Rename moduledef back to xml from bak
                if (isModuleDefInUse)
                {
                    RenameFile(moduleDefPathNew, moduleDefPathOld);
                }
                Thread.Sleep(4000);
                ResetIIS();
                
            }
            else
            {
                logLines.Add("Upgrade not necessary.");
                Console.WriteLine("Upgrade not necessary.");
            }

            //Uncomment to run the BVT here
            //RunBVT(currentBuild);

            //Write log lines to log file
            string[] logLinesArray = logLines.ToArray();
            CreateApplicationLog(logLinesArray);
        }

        public static int GetLastBuildUpgrade()
        {
            int buildNumber = 0;
            //string line;
            string path64 = @"C:\Program Files (x86)\SunView Software\ChangeGear\Server\CGService.exe";
            string path32 = @"C:\Program Files\SunView Software\ChangeGear\Server\CGService.exe";
            //This used to read a file before checking the executable for the version. Now it directly checks the executable.
            //if (File.Exists(@"C:\UpgradeUtil\LastUpgrade.txt"))
            //{
            //    using (StreamReader file = new StreamReader(@"C:\UpgradeUtil\LastUpgrade.txt"))
            //    {
            //        if ((line = file.ReadLine()) != null)
            //            buildNumber = Convert.ToInt32(line);
            //    }
            //    return buildNumber;
            //}
            //else 
            if (File.Exists(path64))
            {
                buildNumber = GetFileVersion(path64);
                return buildNumber;
            }
            else if (File.Exists(path32))
            {
                buildNumber = GetFileVersion(path32);
                return buildNumber;
            }
            else
                return buildNumber;
        }

        public static void SetLastBuildUpgrade(int buildNum)
        {
            DateTime now = DateTime.Now;
            string[] lines = { buildNum.ToString(), now.ToString() };
            System.IO.File.WriteAllLines(@"C:\UpgradeUtil\LastUpgrade.txt", lines);
        }

        public static void CreateApplicationLog(string[] strArray)
        {
            System.IO.File.WriteAllLines(@"C:\UpgradeUtil\ApplicationLog.txt", strArray);
        }

        public static int GetFileVersion(string fPath)
        {
            FileVersionInfo myFileVersionInfo = FileVersionInfo.GetVersionInfo(fPath);
            //string fileVersion = myFileVersionInfo.FileVersion;
            //Console.WriteLine("File: " + myFileVersionInfo.FileDescription + '\n' +
            //    "Version number: " + fileVersion);
            return myFileVersionInfo.FileBuildPart;
        }

        //Close browsers and CG client
        public static void CloseProcesses(Process[] myProcs)
        {
            string tempString;
            foreach (Process thisProcess in myProcs)
            {
                try
                {
                    tempString = string.Format("Closing process: {0}\nID: {1}", thisProcess.ProcessName, thisProcess.Id);
                    logLines.Add(tempString);
                    Console.WriteLine("Closing process: {0}\nID: {1}", thisProcess.ProcessName, thisProcess.Id);
                    thisProcess.Kill();
                }
                catch
                {
                    logLines.Add("Process can't be killed.");
                    Console.WriteLine("Process can't be killed.");
                }
            }
        }

        public static void CloseProcesses()
        {
            Process[] CGprocList = Process.GetProcessesByName("CG");
            Process[] IEprocList = Process.GetProcessesByName("iexplore");
            Process[] FirefoxProcList = Process.GetProcessesByName("firefox");
            Process[] ChromeProcList = Process.GetProcessesByName("chrome");

            if (CGprocList.Length > 0)
                CloseProcesses(CGprocList);
            if (IEprocList.Length > 0)
                CloseProcesses(IEprocList);
            if (FirefoxProcList.Length > 0)
                CloseProcesses(FirefoxProcList);
            if (ChromeProcList.Length > 0)
                CloseProcesses(ChromeProcList);

        }

        public static int GetNewestBuildNumber(string nameOfReleaseFolder)
        {
            string[] directories = Directory.GetDirectories("Z:\\" + nameOfReleaseFolder);
            int direcCount = directories.Length;
            int[] FoldersNum = new int[direcCount];
            int i = 0;
            foreach (string direc in directories)
            {
                Console.WriteLine(direc);
                //Extract build number
                int underScore = 0;
                int indexOfBuild = 0;
                string buildString = "";
                int buildNum = 0;
                for (int ind = 0; ind <= direc.Length; ind++)
                {
                    if (underScore == 3)
                    {
                        indexOfBuild = ind;
                        break;
                    }
                    if (direc[ind] == '_')
                        underScore++;
                }
                buildString = direc.Substring(indexOfBuild, 4);
                buildNum = Convert.ToInt32(buildString);
                FoldersNum[i] = buildNum;
                i++;
            }
            return FoldersNum.Max();

        }

        public static bool DetermineIfUpgradeNecessary(int currentBuild, int newestBuild)
        {
            if (currentBuild < newestBuild)
                return true;
            else
                return false;
        }

        public static string GetNewestBuildPath(string nameOfReleaseFolder, int newestBuildNumber)
        {
            string newestBuildPath = "";
            string[] directories = Directory.GetDirectories("Z:\\" + nameOfReleaseFolder);
            foreach (string direc in directories)
            {
                if (direc.Contains(newestBuildNumber.ToString()))
                {
                    newestBuildPath = string.Format(direc);
                }
            }
            return newestBuildPath;
        }

        public static void CreateInstallerFolder()
        {
            string path =@"C:\UpgradeUtil\Installer";

            if(!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        public static void MapNetworkDrive(string serverPath, string mapArgs)
        {
       
            Process del = Process.Start(@"net.exe", @"use Z: /delete");
         
            Process map = Process.Start(@"net.exe", @"use Z: " + serverPath + " " + mapArgs);
            map.WaitForExit();
            if (System.IO.Directory.Exists(@"Z:\"))
            {
                logLines.Add("Finished mapping network path");
                Console.WriteLine("Finished mapping network path");
            }
            else
            {
                logLines.Add("Network drive not found");
                Console.WriteLine("Network drive not found");
            }
        }

        public static void RenameFile(string oldPath, string newPath)
        {
            if (File.Exists(oldPath))
            {
                try
                {
                    File.Move(oldPath, newPath);
                    logLines.Add(string.Format("File renamed or moved from {0} to {1}", oldPath, newPath));
                    Console.WriteLine(string.Format("File renamed or moved from {0} to {1}", oldPath, newPath));
                }
                catch (Exception)
                {
                    logLines.Add("Rename or move failed!");
                    Console.WriteLine("Rename or move failed!");
                }
            }
        }

        public static string GetChangeGearpath()
        {
            string path32 = @"C:\Program Files\SunView Software\ChangeGear\Server\CGService.exe";
            if (File.Exists(path32))
            {
                return @"C:\Program Files\SunView Software\ChangeGear\";
            }
            else
            {
                return @"C:\Program Files (x86)\SunView Software\ChangeGear\";
            }
        }

        public static void ResetIIS()
        {
            Process restart = new Process();
            restart.StartInfo.Verb = "runas";
            restart.StartInfo.FileName = @"C:\Windows\System32\iisreset.exe";
            restart.Start();
            restart.WaitForExit();
            logLines.Add("IIS server was reset");
            Console.WriteLine("IIS server was reset");
        }

        public static void CreateTRXFolder()
        {
            string path = @"C:\UpgradeUtil\Results\TRX";

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                logLines.Add("TRX folder was created inside the Results folder.");
                Console.WriteLine("TRX folder was created inside the Results folder.");
            }
        }

        public static void RunBVT(int buildNum)
        {
            System.DateTime now = DateTime.Now;
            string nowString = string.Format("{0:yyyy-MM-dd_hhmmtt}", DateTime.Now);
            string nowFileTRX = @"C:\UpgradeUtil\Results\" + nowString + ".trx";
            string nowFileHTM = nowFileTRX + ".htm";
            string trxMoveTo = @"C:\UpgradeUtil\Results\TRX\" + nowString + ".trx";
            string path64 = @"C:\Program Files (x86)\Microsoft Visual Studio 11.0\Common7\IDE\MSTest.exe";
            string path32 = @"C:\Program Files\Microsoft Visual Studio 11.0\Common7\IDE\MSTest.exe";
            string cmd;

            //Create necessary folders and files
            CreateTRXFolder();
            CreateBVTFileDependencies();
            

            //Commands to run BVT
            if (File.Exists(path64))
            {
                cmd = @"
cd C:\Program Files (x86)\Microsoft Visual Studio 11.0\Common7\IDE
MSTest /testcontainer:""C:\CGSpecFlowBVT\bin\Debug\CGSpecFlowBVT.dll"" /resultsfile:""" +
    nowFileTRX + "\"";
            }
            else if (File.Exists(path32))
            {
                cmd = @"
cd C:\Program Files\Microsoft Visual Studio 11.0\Common7\IDE
MSTest /testcontainer:""C:\CGSpecFlowBVT\bin\Debug\CGSpecFlowBVT.dll"" /resultsfile:""" +
    nowFileTRX + "\"";
            }
            else
            {
                logLines.Add("MSTest not found. Aborting execution.");
                Console.WriteLine("MSTest not found. Aborting execution.");
                return;
            }

            //Commands to send emails
            string BLATcmd = @"
cd C:\UpgradeUtil\Blat
blat -to msuarin@sunviewsoftware.com,mwyman@sunviewsoftware.com,klettiero@sunviewsoftware.com,tluu@sunviewsoftware.com -server 192.168.1.203 -f msuarin@sunviewsoftware.com -subject ""Daily ChangeGear BVT Test Results - " +
buildNum + @""" -body ""The ChangeGear BVT has been run. See the attachment for the test results."" -attach """ +
nowFileHTM + "\"";

            string BVTbat = @"C:\UpgradeUtil\BVT.bat";
            File.WriteAllText(BVTbat, cmd);
            logLines.Add("BVT batch file created.");
            Console.WriteLine("BVT batch file created.");

            string BLATbat = @"C:\UpgradeUtil\blat.bat";
            File.WriteAllText(BLATbat, BLATcmd);
            logLines.Add("Blat batch file created.");
            Console.WriteLine("Blat batch file created");

            //Run the BVT batch file
            logLines.Add("Running the BVT batch file. Selenium tests are about to run shortly...");
            Console.WriteLine("Running the BVT batch file. Selenium tests are about to run shortly...");
            Process runBat = Process.Start(BVTbat);
            runBat.WaitForExit();

            //Wait for 10 seconds
            Thread.Sleep(10000);

            //Convert TRX to HTM
            ProcessStartInfo convertToHTMProcess = new ProcessStartInfo();
            convertToHTMProcess.FileName = @"C:\UpgradeUtil\Results\trx2html.exe";
            convertToHTMProcess.Arguments = nowFileTRX;
            Process runConvert = Process.Start(convertToHTMProcess);
            runConvert.WaitForExit();
            logLines.Add("Created an htm version of the results.");
            Console.WriteLine("Created an htm version of the results.");

            //Wait for 10 seconds
            Thread.Sleep(10000);

            //Move the trx file to the TRX folder
            RenameFile(nowFileTRX, trxMoveTo);

            //Run the blat batch file to send the email
            logLines.Add("Running the blat batch file. Emails are to be sent shortly...");
            Console.WriteLine("Running the blat batch file. Emails are to be sent shortly...");
            Process runBlatBat = Process.Start(BLATbat);
            runBlatBat.WaitForExit();
        }

        public static void CreateBVTFileDependencies()
        {
            if (!File.Exists(@"C:\QA_BVT_Files\test1.txt"))
            {
                Directory.CreateDirectory(@"C:\QA_BVT_Files");
                using (File.Create(@"C:\QA_BVT_Files\test1.txt")) { };
                logLines.Add("BVT attachment file created");
                Console.WriteLine("BVT attachment file created");
            }
        }
    }
}
