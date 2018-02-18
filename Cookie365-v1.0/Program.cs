/*  
    (c) Copyright 2014-2018 Fabio Cuneaz 

    This file is part of Cookie365.

    Cookie365 is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    Cookie365 is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with Cookie365.  If not, see <http://www.gnu.org/licenses/>.

*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Net;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;


namespace Cookie365
{
    class Program
    {
        // Import DLL to set Cookies in IE
        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetSetCookie(string lpszUrlName, string lpszCookieName, string lpszCookieData);
        
        static void Main(string[] args)
        {
            // Set Default Argument values
            bool quiet = false;
            bool debug = false;
            bool useIntegratedAuth = true;
            bool mount = false;
            string sharepointUrl = null;
            Uri sharepointUri = null;
            string username = null;
            string password = null;
            string disk = null;
            double expire = 0;
            string homedir = "";

            // Create instance of Log to display messages
            Log log = new Log();

            //Parse args
            Args CommandLine = new Args(args);
            if (CommandLine["s"] == null)
             log.WriteLine("Error: SharePoint URL not specified !\n\nUsage: Cookie365 -s URL [-u user@domain.com | -d domain.com] [-p {password}] [-quiet] [-mount [disk] [-homedir]] [-debug]");
            else
            {
                try
                {
                    // Retrieve SharePoint URL and Create URI
                    sharepointUrl = CommandLine["s"];
                    sharepointUri = new Uri(sharepointUrl);

                    if (!quiet)
                    {
                        // Message
                        log.WriteLine("============= Cookie365 v0.9 - (C)opyright 2014-2018 Fabio Cuneaz =============\n");
                        log.WriteLine("OneDrive for business URL", sharepointUrl);
                    }

                    // If username is specified use it, otherwise try to get the user User Principal Name from AD
                    if ((username=CommandLine["u"]) == null) 
                    {
                        // Try to get User Principal Name from Active Directory
                        username = System.DirectoryServices.AccountManagement.UserPrincipal.Current.UserPrincipalName;
                        if (username != null) { if (CommandLine["d"] != null) username = username.Split('@')[0] + "@" + CommandLine["d"]; }
                        else throw new Exception("Cannot retrieve your username from Active Directory. Username cannot be empty. Use -u option to specify a username");
                    }

                    // If password is specified, use it, otherwise try integrated authentication
                    if (CommandLine["p"] != null) { password = CommandLine["p"]; useIntegratedAuth = false; }

                    // Set the flag for quiet mode
                    if (CommandLine["quiet"] != null) 
                    {
                        quiet = true;
                    }
                    else if (CommandLine["debug"] != null) { debug = true; }    

                    // If asked to mount sharepoint as a share
                    disk = CommandLine["mount"];
                    if (disk != null)
                    {
                        mount = true;
                        if (disk == "true") disk = "*";
                    }            

                    if (CommandLine["offset"] != null)
                    {
                       expire = Convert.ToDouble(CommandLine["expire"]);
                    }
                    else
                    {
                        TimeZone tz = TimeZone.CurrentTimeZone;
                        TimeSpan ts = tz.GetUtcOffset(DateTime.Now);
                        expire = - (double)ts.TotalMinutes;
                    }

                    if (debug) log.Dbg("Offset from GMT for cookie expiration", expire.ToString());

                    if (CommandLine["homedir"] != null)
                    {
                        String user = username.Split('@')[0];
                        String domain = username.Split('@')[1];
                        homedir = "DavWWWRoot\\personal\\" + user + "_" + domain.Split('.')[0] + "_" + domain.Split('.')[1] + "\\Documents";
                    }

                    // if not quiet, display parameters
                    if(!quiet)
                    {
                        log.WriteLine("User", username);
                        log.WriteLine("Use Windows Integrated Authentication", useIntegratedAuth.ToString());
                        if (homedir != "") log.WriteLine("Home Directory", homedir);
                        if (mount) log.WriteLine("Mount as disk", disk);
                    }

                    // Run Asynchronously and wait for cookie retrieval
                    RunAsync(sharepointUri, username, password, useIntegratedAuth, !quiet, debug).Wait();
                    
                    // If
                    if (SpoAuthUtility.Current != null)
                    {
                        if (!quiet) Console.Write("Setting Cookies in OS...");
                        try
                        {
                            // Create the cookie collection object for sharepoint URI
                            CookieCollection cookies = SpoAuthUtility.Current.cookieContainer.GetCookies(sharepointUri);

                            // Extract the base URL in case the URL provided contains nested paths (e.g. https://contoso.sharepoint.com/abc/ddd/eed)
                            // The cookie has to be set for the domain (contoso.sharepoint.com), otherwise it will not work
                            String baseUrl = sharepointUri.Scheme + "://" + sharepointUri.Host;

                            if (InternetSetCookie(baseUrl, null, cookies["FedAuth"].ToString() + "; Expires = " + cookies["FedAuth"].Expires.AddMinutes(expire).ToString("R")))
                            {
                                if (InternetSetCookie(baseUrl, null, cookies["rtFA"].ToString() + "; Expires = " + cookies["rtFA"].Expires.AddMinutes(expire).ToString("R")))
                                {
                                    if (!quiet) log.WriteLine("[OK]");
                                    if (debug) log.Dbg("Cookie Expiration", cookies["FedAuth"].Expires.AddMinutes(expire).ToString("R"));
                                  if (mount)
                                  {
                                      try
                                      {
                                          String cmdArgs = "/c net use " + disk + " \\\\" + sharepointUri.Host + "@ssl" + sharepointUri.PathAndQuery.Replace("/", "\\") + homedir;
                                          if (!quiet) Console.Write("Mounting Share...");
                                          System.Diagnostics.Process Process = new System.Diagnostics.Process();
                                          Process.StartInfo = new System.Diagnostics.ProcessStartInfo("cmd", cmdArgs);
                                          Process.StartInfo.RedirectStandardOutput = true;
                                          Process.StartInfo.UseShellExecute = false;
                                          Process.Start();
                                          Process.WaitForExit();
                                          String mountOutput = Process.StandardOutput.ReadToEnd();
                                          if (!quiet)
                                          {
                                              log.WriteLine("[OK]");
                                              log.WriteLine(mountOutput);
                                          }
                                      }
                                      catch (Exception e)
                                      { log.Err("Mounting Share", e.Message); }
                                  }
                                }
                            }
                        }
                        catch (Exception e)
                        { log.Err("setting Cookies", e.Message); }

                    }
                }
                catch (Exception e)
                { log.Err(e.Message); }
            }
        }

        static async Task RunAsync(Uri sharepointUri, string username, string password, bool useIntegratedAuth, bool verbose, bool debug)
        {
            await SpoAuthUtility.Create(sharepointUri, username, password, useIntegratedAuth, verbose, debug);
        }
    }
}
