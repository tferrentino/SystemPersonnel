using System;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;

namespace LoadData
{
    public class RecursiveFileProcessor
    {
        public static void Main(string[] args)
        {
            foreach (string path in args)
            {
                        if (File.Exists(path))
                {
                    // Process all files listed in a particular folder. 
                    // Currently the ProcessFile method is initially looking in 
                    // the C:\CS\TestIN folder
                    // The variable "path" is the absolute path to files within the folder.
                    // 
                    ProcessFile(path);
                }
                else if (Directory.Exists(path))
                {
                    // This path is a directory
                    // Process all files in all subfolders in a directory tree. 
                    // Currently the ProcessDirectory method is initially looking in 
                    // the C:\CS parent folder
                    // The variable "path" is the absolute path to files within the folder.
                    // 
                    ProcessDirectory(path);
                }
                else
                {
                    Console.WriteLine("{0} is not a valid file or directory.", path);
                }
            }
            Console.WriteLine("End of Directory");
            Console.ReadKey();
        }


        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        // This path is a directory
        // Process all files in all subfolders in a directory tree. 
        // Currently the ProcessDirectory method is initially looking in 
        // the C:\CS parent folder
        // The variable "path" is the absolute path to files within the folder.
        // 
        public static async Task ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                if (Path.GetDirectoryName(fileName) != @"C:\CS\Completed")
                {
                    ProcessFile(fileName);
                }
                    

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                await ProcessDirectory(subdirectory);
        }

        // Process all files listed in a particular folder. 
        // Currently the ProcessFile method is initially looking in 
        // the C:\CS\TestIN folder
        // The variable "path" is the absolute path to files within the folder.
        // 
        private static void ProcessFile(string path)
        {
            string extension = Path.GetExtension(path);
            string fileName = "";
            // The "Completed" subfolder will hold all files that have been processed (Migrated to SQL Server)
            string completed = @"C:\CS\Completed\";
            string fileNameOnly = Path.GetFileNameWithoutExtension(path);
            UpdateVariables updateXml = new UpdateVariables();

            // string completed = "C:\\CS\\Completed\\";
            if (path.Length > path.LastIndexOf(@"\", StringComparison.Ordinal) + 1)
            {
                fileName = path.Substring(path.LastIndexOf(@"\", StringComparison.Ordinal) + 1);
            }
            //
            // Check if file is a valid Access database. 
            // If not then go to the next file in the list
            // Else process the Access database.
            //
            if (extension == ".mdb" || extension == ".accdb" || extension == ".MDB" || extension == ".ACCDB")
            {
                completed += @"\" + fileName;
                //
                //  Process the Access dataabase
                //
                updateXml.MigrateFile(fileNameOnly);
                if (File.Exists(completed))
                {
                    File.Delete(completed);
                    File.Move(path, completed);
                    Console.WriteLine("Moved and Replaced file '{0}' with {1}", path, completed);
                }
                else
                {
                    File.Move(path, completed);
                    Console.WriteLine("Moved file '{0}' to {1}", path, completed);
                }
                Updatempclients(fileNameOnly);
                Console.WriteLine("Processed file '{0}'.", path);
            }
        }

        private static void Updatempclients(string fileNameOnly)
        {
            try
            {
                // *******************************************************
                // Tom Ferrentino
                // Update mpclients with new client database information
                // July 21, 2017
                // *******************************************************
                using (SqlConnection connection = new SqlConnection(@"Server=SQLDEV\DEVSQL2014;Integrated Security=SSPI;Initial Catalog=Micropay_MPCL"))
                {
                    string strSQL;
                    /*
                    strSQL = @"UPDATE dbo.mpclients  SET clientID = '" + Right(fileNameOnly, 3) + "', ";
                    strSQL += @"[client name] = '" + fileNameOnly + "', ";
                    strSQL += @"DataSource = 'SQLDEV\DEVSQL2014', initialCatalog = 'Micropay_" + fileNameOnly + "', ";
                    strSQL += @"UserName = 'Micropay_User=', Password='M1cr0p@y'";
                    */
                    strSQL = @"INSERT INTO dbo.mpclients (clientID, [client name], DataSource, initialCatalog, Username, Password) ";
                    strSQL += @"VALUES ('" + Right(fileNameOnly, 3) + "', '" + fileNameOnly + @"', 'SQLDEV\DEVSQL2014', 'Micropay_" + fileNameOnly + "', 'Micropay_User', 'M1cr0p@y')";
                    connection.Open();
                    SqlCommand updateCommand = new SqlCommand(strSQL, connection);
                    updateCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error updating mpclients Database = {0}, Error: {1}", fileNameOnly, ex);
                throw;
            }
            finally
            {
                Console.WriteLine("Successfully Updated mpclients for {0}", fileNameOnly);
            }
        }

        private static string Right(string value, int length)
        {
            return value.Substring(value.Length - length);
        }
       
    }
}
