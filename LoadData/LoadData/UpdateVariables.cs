using System;
using System.Diagnostics;
using System.Text;
using System.Xml.Linq;

namespace LoadData
{
    class UpdateVariables
    {
        /*  Name:       MigrateFile
         *  Author:     Tom Ferrentino 
         *  Date:       June 17, 2017
         *  Description: This method prepares the Access97 database
         *  to be migrated to SQL Server 2014. The following proceses need to 
         *  occur for this to happen:
         *  1. Modify the SQL Server Migration Assistant (SSMA) XML scripts
         *  2. Build the XML script arguement string
         *  3. Pass the XML scripts as arguements to the SSMA command line tool
         *  4. Call the "RunExternalExe" method which starts the SSMA command line process
         */
        public void MigrateFile(string fileName)
        {
            // ReSharper disable once SuggestVarOrType_SimpleTypes
            DateTime moment = DateTime.Today;
            int year = moment.Year;
            int month = moment.Month;
            int day = moment.Day;


            XDocument xdoc = XDocument.Load(@"C:\CS\DB\Schemas\Variables.xml");
            foreach (XElement element in xdoc.Descendants())
            {
                Console.WriteLine(element.FirstAttribute.Value + " " + element.LastAttribute.Value);
                if (element.FirstAttribute.Value.Contains("AccessDatabase") &&
                    !element.FirstAttribute.Value.Contains("AccessDatabaseFile"))
                {
                    element.LastAttribute.Value = fileName;
                    Console.WriteLine("Updated Access Database Name");
                }
                if (element.FirstAttribute.Value.Contains("SQLServerDb"))
                {
                    element.LastAttribute.Value = "Micropay_MPM" + Right(fileName, 3);
                    Console.WriteLine("Updated SQL Server Databasa Name");
                }
                if (element.FirstAttribute.Value.Contains("project_name"))
                {
                    element.LastAttribute.Value = "Migrate_" + year.ToString() + month.ToString() + day.ToString() + Right(fileName, 3);
                    Console.WriteLine("Updated Project Folder Name");
                }
            }
            xdoc.Save(@"C:\CS\DB\Schemas\Variables.xml");
            Console.ReadKey();
            string xmlArgs = " -s ";
            xmlArgs += @"""C:\CS\DB\Schemas\Script.xml""" + " -v ";
            xmlArgs += @"""C:\CS\DB\Schemas\Variables.xml""" + " -c ";
            xmlArgs += @"""C:\CS\DB\Schemas\Server.xml""";
            xmlArgs += @" -l c:\CS\LogFile_" + Right(fileName, 3) + ".txt";
            var retOutput = RunExternalExe("SSMAforAccessConsole32", xmlArgs);
            Console.WriteLine("Migration Results for " + fileName + ": " + retOutput);
        }

        private static string Right(string value, int length)
        {
            return value.Substring(value.Length - length);
        }
        /*  Name:           RunExternalExe
         *  Author:         Tom Ferrentino 
         *  Date:           June 17, 2017
         *  Description:    This method kicks off the SSMA command line 
         *                  that migrates the Access97 client to a new 
         *                  SQL Server 2014 database client.
         *  
         */
        public string RunExternalExe(string filename, string arguments = null)
        {
            Process process = new Process();
            process.StartInfo.FileName = filename;

            if (!string.IsNullOrEmpty(arguments))
            {
                process.StartInfo.Arguments = arguments;
            }
            
            process.StartInfo.CreateNoWindow = true;
            // process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.StartInfo.Verb = "runas";
            process.StartInfo.UseShellExecute = false;

            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.RedirectStandardOutput = true;
            var stdOutput = new StringBuilder();
            process.OutputDataReceived += (sender, args) =>
            {
                if (sender == null) throw new ArgumentNullException(nameof(sender));
                stdOutput.AppendLine(args.Data);
            }; // Use AppendLine rather than Append since args.Data is one line of output, not including the newline character.

            string stdError = null;
            try
            {
                process.Start();
                process.BeginOutputReadLine();
                stdError = process.StandardError.ReadToEnd();
                process.WaitForExit();
                // Call Load Scripts Procedure in Sequence //
            }
            catch (Exception e)
            {
                throw new Exception("OS error while executing " + Format(filename, arguments) + ": " + e.Message, e);   
            }

            if (process.ExitCode == 0)
            {
                return stdOutput.ToString();
            }
            else
            {
                var message = new StringBuilder();

                if (!string.IsNullOrEmpty(stdError))
                {
                    message.AppendLine(stdError);
                }

                if (stdOutput.Length != 0)
                {
                    message.AppendLine("Std output:");
                    message.AppendLine(stdOutput.ToString());
                }

                throw new Exception(Format(filename, arguments) + " finished with exit code = " + process.ExitCode + ": " + message);
            }
        }

        private static string Format(string filename, string arguments)
        {
            return "'" + filename +
                ((string.IsNullOrEmpty(arguments)) ? string.Empty : " " + arguments) +
                "'";
        }


    }
}
