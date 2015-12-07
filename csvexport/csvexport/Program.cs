//Copyright 2015, Combined Public Communications
/*
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

using Renci.SshNet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace InmateExport
{
    class Program
    {
        //settings
        static void Main()
        {
            var s = new Settings();
            var log = new Log();

            try
            {
                //fill datatable
                Console.Write("Filling data table...");

                var sqlConnection = new SqlConnection(s.db.connString);
                var sQuery = new StringBuilder();

                sQuery.Append("SELECT * FROM ").Append(s.db.view).Append(" ");

                if (s.db.filter != null && s.db.filter != "")
                    sQuery.Append("WHERE ").Append(s.db.filter).Append(" ");
                if (s.db.sort != null && s.db.sort != "")
                    sQuery.Append("ORDER BY ").Append(s.db.sort);

                var sqlCmd = new SqlCommand(sQuery.ToString(), sqlConnection);
                var sqlDa = new SqlDataAdapter(sqlCmd);
                var dt = new DataTable();

                sqlConnection.Open();
                sqlDa.Fill(dt);
                sqlConnection.Close();

                Console.WriteLine("DONE");

                //convert data table to delimited value file
                if (dt.Rows.Count <= 0)
                    throw new Exception("No data to write.");

                Console.Write("Converting to file...");

                var sData = new StringBuilder();

                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < row.ItemArray.Length; i++)
                    {
                        sData.Append(row[i].ToString());
                        if (i != row.ItemArray.Length - 1)
                            sData.Append(s.general.delimiter);
                    }

                    sData.AppendLine();
                }

                //write to file
                Console.Write("Wiriting to file...");
                var sFile = new StringBuilder();
                sFile
                    .Append(s.general.fileName)
                    .Append(DateTime.Now.Month.ToString("d2"))
                    .Append(DateTime.Now.Day.ToString("d2"))
                    .Append(DateTime.Now.Year.ToString("d4"))
                    .Append(DateTime.Now.Hour.ToString("d2"))
                    .Append(DateTime.Now.Minute.ToString("d2"))
                    .Append(DateTime.Now.Second.ToString("d2"))
                    .Append(s.general.fileExt);

                var fullPath = Environment.CurrentDirectory + "/" + sFile.ToString();

                using (var sw = File.CreateText(fullPath))
                    sw.Write(sData.ToString());

                Console.WriteLine("DONE");
                log.Info("New export file created.");

                //copy file to ftp
                if (s.general.useFTP)
                {
                    Console.Write("Copying file to SFTP server...");
                    var ss = new SFTPSettings();

                    AuthenticationMethod method =
                        new PasswordAuthenticationMethod(ss.user, ss.pass);
                    ConnectionInfo info =
                        new ConnectionInfo(ss.server, ss.port, ss.user, method);
                    SftpClient client = new SftpClient(info);

                    client.Connect();

                    using (var stream = File.Open(fullPath, FileMode.Open))
                        client.UploadFile(stream, ss.dir + sFile.ToString(), null);

                    client.Disconnect();

                    Console.WriteLine("DONE");
                    log.Info("File copied to SFTP server.");
                }

                //copy file to local directory
                if (s.general.useLocal)
                {
                    Console.Write("Copying to local directory...");
                    if (!Directory.Exists(s.general.localDir))
                        Directory.CreateDirectory(s.general.localDir);
                    File.Copy(fullPath, s.general.localDir + sFile.ToString());
                    Console.WriteLine("DONE");
                    log.Info("File copied to local directory.");
                }

                Console.WriteLine("Task Complete.");
                log.Info("Task Complete.");

                //delete temp file
                Console.WriteLine("Deleting temp file...");
                if (File.Exists(fullPath))
                    File.Delete(fullPath);
            }
            catch (Exception ex)
            {
                log.Error("File export failed.", ex);

                //send error email if enabled
                if (s.email.enabled)
                {
                    var sBody = new StringBuilder();
                    sBody
                        .AppendLine("File export has failed. Please see error below.")
                        .AppendLine(ex.ToString());

                    MailMessage message = new MailMessage();
                    message.To.Add(s.email.to);
                    message.Subject = "File Export Failed";
                    message.From = new MailAddress(s.email.from);
                    message.Body = sBody.ToString();
                    message.IsBodyHtml = true;

                    if (s.email.cc != null && s.email.cc != "")
                        message.CC.Add(s.email.cc);
                    if (s.email.bcc != null && s.email.bcc != "")
                        message.Bcc.Add(s.email.bcc);

                    SmtpClient smtp = new SmtpClient(s.email.server);

                    if (s.email.authReq)
                    {
                        smtp.UseDefaultCredentials = false;
                        smtp.Credentials = new System.Net.NetworkCredential(s.email.user, s.email.pass);
                    }
                    else
                    {
                        smtp.UseDefaultCredentials = true;
                    }

                    smtp.Send(message);
                    smtp.Dispose();
                }

                Console.WriteLine("FAILED");
                Console.WriteLine();
                Console.WriteLine("CHECK LOG FILE FOR DETAILS");
                Console.WriteLine("Press any key to stop the program...");
                Console.ReadKey();
            }
        }
    }
}

#region settings
public class Settings
{
    public GeneralSettings general { get; set; }
    public DatabaseSettings db { get; set; }
    public EmailSettings email { get; set; }

    public Settings()
    {
        general = new GeneralSettings();
        db = new DatabaseSettings();
        email = new EmailSettings();
    }
}
public class GeneralSettings
{
    public bool useFTP { get; set; }
    public bool useLocal { get; set; }
    public string localDir { get; set; }
    public char delimiter { get; set; }
    public string fileExt { get; set; }
    public string fileName { get; set; }

    public GeneralSettings()
    {
        var s = ConfigurationManager.GetSection("general") as NameValueCollection;
        useFTP = Convert.ToBoolean(s["useFTP"]);
        useLocal = Convert.ToBoolean(s["useLocalDirectory"]);
        localDir = s["localDirectory"];
        delimiter = Convert.ToChar(s["delimiter"]);
        fileExt = s["fileExt"];
        fileName = s["fileName"];
    }
}

public class LogSettings
{
    public bool enable { get; set; }
    public string dir { get; set; }
    public string filename { get; set; }

    public LogSettings()
    {
        var s = ConfigurationManager.GetSection("log") as NameValueCollection;
        enable = Convert.ToBoolean(s["enable"]);
        dir = s["directory"];
        filename = s["fileName"];
    }
}

public class DatabaseSettings
{
    public string connString { get; set; }
    public string view { get; set; }
    public string filter { get; set; }
    public string sort { get; set; }

    public DatabaseSettings()
    {
        var s = ConfigurationManager.GetSection("database") as NameValueCollection;
        connString = s["connectionString"];
        view = s["view"];
        filter = s["filter"];
        sort = s["sort"];
    }
}

public class SFTPSettings
{
    public string server { get; set; }
    public int port { get; set; }
    public string dir { get; set; }
    public string user { get; set; }
    public string pass { get; set; }

    public SFTPSettings()
    {
        var s = ConfigurationManager.GetSection("sftp") as NameValueCollection;
        server = s["server"];
        port = Convert.ToInt16(s["port"]);
        dir = s["uploadDir"];
        user = s["userName"];
        pass = s["password"];
    }
}

public class EmailSettings
{
    public bool enabled { get; set; }
    public bool authReq { get; set; }
    public string server { get; set; }
    public string user { get; set; }
    public string pass { get; set; }
    public string from { get; set; }
    public string to { get; set; }
    public string cc { get; set; }
    public string bcc { get; set; }

    public EmailSettings()
    {
        var s = ConfigurationManager.GetSection("email") as NameValueCollection;
        enabled = Convert.ToBoolean(s["enabled"]);
        authReq = Convert.ToBoolean(s["authReq"]);
        server = s["server"];
        user = s["userName"];
        pass = s["passWord"];
        from = s["from"];
        to = s["to"];
        cc = s["cc"];
        bcc = s["bcc"];
    }
}
#endregion

#region log
public class Log
{
    LogSettings ls = new LogSettings();

    public void Info(string message)
    {
        var sb = new StringBuilder();
        sb
            .Append(GetDate())
            .Append(" ~ INFO: ").Append(message)
            .AppendLine();

        WriteLog(sb.ToString());
    }

    public void Error(string message, Exception ex = null)
    {
        var sb = new StringBuilder();
        sb.Append(GetDate()).Append(" ~ ERROR:").Append(message).AppendLine();

        if (ex != null)
        {
            sb.Append("Source: ").Append(ex.Source).AppendLine()
              .Append("Error #: ").Append(ex.HResult).AppendLine()
              .Append("Message: ").Append(ex.Message).AppendLine()
              .Append("Inner Exception: ").Append(ex.InnerException).AppendLine()
              .Append("Details: ").AppendLine();

            foreach (DictionaryEntry d in ex.Data)
            {
                sb.Append("KEY: ").Append(d.Key.ToString()).AppendLine()
                  .Append("VALUE: ").Append(d.Value.ToString());
            }
        }

        WriteLog(sb.ToString());
    }

    private void WriteLog(string message)
    {
        if (ls.enable)
        {
            if (!Directory.Exists(ls.dir))
                Directory.CreateDirectory(ls.dir);

            var date = DateTime.Now;
            var path = string.Format(@"{0}{1}", ls.dir, GetLatestFile());

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                    sw.Write(message);
            }
            else
            {
                var info = new FileInfo(path);
                using (StreamWriter sw = File.AppendText(path))
                    sw.Write(message);
            }
        }
    }

    private string GetLatestFile()
    {
        var di = new DirectoryInfo(ls.dir);
        FileInfo file = di.GetFiles().Where(d => d.Extension == ".log").OrderByDescending(f => f.LastWriteTime).FirstOrDefault();

        if (file == null)
            return ls.filename + ".log";
        else if (file.Length >= 5242880)
            return ls.filename + di.GetFiles().Where(d => d.Extension == ".log").Count() + ".log";
        else
            return file.Name;
    }

    private static string GetDate()
    {
        return string.Format("{0} {1}",
            DateTime.Now.ToShortDateString(),
            DateTime.Now.ToShortTimeString());
    }
}
#endregion
