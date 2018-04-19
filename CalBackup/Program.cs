using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;


namespace CalBackup
{
    class Program
    {
        public static ExchangeService exService;
        public static Folder fldCalBackup = null;
        public static CalendarFolder fldCalRestore = null;
        public static string strClientID = "535ee24f-50a9-41e5-ba24-acbb4a44a662"; 
        public static string strRedirURI = "https://CalBackup";
        public static string strAuthCommon = "https://login.microsoftonline.com/common";
        public static string strSrvURI = "https://outlook.office365.com";
        public static string strDisplayName = "";
        public static string strSMTPAddr = "";
        public static bool bRestore = false;
        public static int cItems = 0;


        static void Main(string[] args)
        {
            string strAcct = "";
            string strTenant = "";
            string strEmailAddr = "";
            bool bMailbox = false;
            NameResolutionCollection ncCol = null;
            Folder fldCal = null;
            
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].ToUpper() == "-M" || args[i].ToUpper() == "/M") // mailbox mode - use impersonation to get to another mailbox
                    {
                        if (args[i + 1].Length > 0)
                        {
                            strEmailAddr = args[i + 1];
                            bMailbox = true;
                        }
                        else
                        {
                            Console.WriteLine("Please enter a valid SMTP address for the mailbox.");
                            ShowHelp();
                            return;
                        }
                    }
                    if (args[i].ToUpper() == "-R" || args[i].ToUpper() == "/R")
                    {
                        bRestore = true;
                    }
                    if (args[i].ToUpper() == "-?" || args[i].ToUpper() == "/?") // display command switch help
                    {
                        ShowInfo();
                        ShowHelp();
                        return;
                    }
                }
            }

            ShowInfo();

            exService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            exService.UseDefaultCredentials = false;

            Console.Write("Press <ENTER> to enter credentials.");
            Console.ReadLine();
            Console.WriteLine();

            AuthenticationResult authResult = GetToken();
            if (authResult != null)
            {
                exService.Credentials = new OAuthCredentials(authResult.AccessToken);
                strAcct = authResult.UserInfo.DisplayableId;
            }
            else
            {
                return;
            }
            strTenant = strAcct.Split('@')[1];
            exService.Url = new Uri(strSrvURI + "/ews/exchange.asmx");

            if (bMailbox)
            {
                ncCol = DoResolveName(strEmailAddr);
                exService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, strEmailAddr);
            }
            else
            {
                ncCol = DoResolveName(strAcct);
            }

            if (ncCol == null)
            {
                // Didn't get a NameResCollection, so error out.
                Console.WriteLine("");
                Console.WriteLine("Exiting the program.");
                return;
            }

            if (ncCol[0].Contact != null)
            {
                strDisplayName = ncCol[0].Contact.DisplayName;
                strEmailAddr = ncCol[0].Mailbox.Address;
                if (!bRestore)
                {
                    Console.WriteLine("Backing up Calendar for " + strDisplayName);
                }
                else
                {
                    Console.WriteLine("Restoring Calendar for " + strDisplayName);
                }
                
            }
            else
            {
                if (!bRestore)
                {
                    Console.WriteLine("Backing up Calendar for " + strAcct);
                }
                else
                {
                    Console.WriteLine("Restoring Calendar for " + strAcct);
                }

            }

            // Get or create the backup folder...
            GetBackupFld();

            // Get or create the Restore folder
            if (bRestore)
            {
                GetRestoreFld();
            }
            else
            {
                // Get the Calendar folder
                try
                {
                    Console.WriteLine("Connecting to the Calendar.");
                    fldCal = Folder.Bind(exService, WellKnownFolderName.Calendar, new PropertySet(PropertySet.IdOnly));
                }
                catch (ServiceResponseException ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("");
                    Console.WriteLine("Could not connect to this user's mailbox or calendar.");
                    Console.WriteLine(ex.Message);
                    Console.ResetColor();
                    return;
                }
            }

            if (!bRestore)
            {
                // now go get the items and copy them to the backup folder.
                Console.WriteLine("Copying all items from Calendar to the CalBackup folder.");
                CopyItems(fldCal, fldCalBackup);

                Console.WriteLine("\r\n");
                Console.WriteLine("===============================================================");
                Console.WriteLine("Copied " + cItems.ToString() + " items to the CalBackup folder.");
                Console.WriteLine("===============================================================");
            }
            else
            {
                Console.WriteLine("Restoring items to the CalRestore folder.");
                CopyItems(fldCalBackup, fldCalRestore);

                Console.WriteLine("\r\n");
                Console.WriteLine("==================================================================");
                Console.WriteLine("Restored " + cItems.ToString() + " items to the CalRestore folder.");
                Console.WriteLine("==================================================================");
            }

            //Console.Write("Press a key to exit");
            //Console.Read();
            return;
        }

        public static void ShowInfo()
        {
            Console.WriteLine("");
            Console.WriteLine("=========");
            Console.WriteLine("CalBackup");
            Console.WriteLine("=========");
            Console.WriteLine("Backs up Calendar Items to the CalBackup subfolder under the Recoverable Items folder.\r\n");
        }

        public static void ShowHelp()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("CalBackup [-M <SMTP Address>] [-R] [-?]");
            Console.WriteLine("");
            Console.WriteLine("-M   [Mailbox - will connect to the mailbox and perform the backup.]");
            Console.WriteLine("-R   [Restores items from the CalBackup folder to the CalRestore folder.]");
            Console.WriteLine("-?   [Shows this usage information.]");
            Console.WriteLine("");
        }

        // Perform copying items from one folder to another
        public static void CopyItems(Folder fldSrc, Folder fldDst)
        {
            bool bMore = true;
            int iOffset = 0;
            int iPageSize = 500;
            FindItemsResults<Item> findResults = null;
            ItemView cView = new ItemView(iPageSize, iOffset, OffsetBasePoint.Beginning);
            cView.OrderBy.Add(ItemSchema.LastModifiedTime, SortDirection.Descending);
            char[] cSpin = new char[] { '/', '-', '\\', '|' };

            while (bMore)
            {
                int i = 0;
                int n = 0;
                findResults = fldSrc.FindItems(cView);

                if (!(findResults.Items.Count == 0))
                {
                    foreach (Item item in findResults.Items)
                    {
                        i++;
                        if (i % 5 == 0)
                        {
                            Console.SetCursorPosition(0, Console.CursorTop);
                            Console.Write("");
                            Console.Write(cSpin[n % 4]);
                            n++;
                        }
                        item.Copy(fldDst.Id);
                        cItems++;
                    }

                    bMore = findResults.MoreAvailable;
                    if (bMore)
                    {
                        cView.Offset += iPageSize;
                    }
                }
                else
                {
                    Console.WriteLine("The source folder " + fldSrc.DisplayName + " is empty. Nothing to copy.");
                }
            }
        }

        // If not there - then it will create it. Otherwise finds it.
        public static void GetBackupFld()
        {
            Folder fld = new CalendarFolder(exService);
            fld.DisplayName = "CalBackup";
             
            bool bEmpty = false;

            if (!bRestore)
            {
                try
                {
                    Console.WriteLine("Accessing/Creating the CalBackup folder.");
                    fld.Save(WellKnownFolderName.RecoverableItemsRoot);
                }
                catch (ServiceResponseException ex)
                {
                    if (!(ex.ErrorCode.ToString().ToUpper() == "ERRORFOLDEREXISTS"))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("");
                        Console.WriteLine(ex.Message);
                        Console.ResetColor();
                        return;
                    }
                    else
                    {
                        fld = FindFolder("CalBackup");
                        Console.WriteLine("Removing previously backed up items from the CalBackup folder.");
                        bEmpty = EmptyFolder(fld);
                        Console.WriteLine("\r\nFinished removing items from the CalBackup folder.");
                    }
                }
                fldCalBackup = fld;
            }
            else
            {
                Console.WriteLine("Accessing/Creating the CalBackup folder.");
                fldCalBackup = FindFolder("CalBackup");
            }
        }

        // If not there - then it will create it. Otherwise finds it.
        public static void GetRestoreFld()
        {
            CalendarFolder fld = new CalendarFolder(exService);
            fld.DisplayName = "CalRestore";

            bool bEmpty = false;

            try
            {
                Console.WriteLine("Accessing/Creating the CalRestore folder.");
                fld.Save(WellKnownFolderName.Calendar);
            }
            catch (ServiceResponseException ex)
            {
                if (!(ex.ErrorCode.ToString().ToUpper() == "ERRORFOLDEREXISTS"))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("");
                    Console.WriteLine(ex.Message);
                    Console.ResetColor();
                    return;
                }
                else
                {
                    fld = FindCalFolder("CalRestore");
                    Console.WriteLine("Removing previously restored items from the CalRestore folder.");
                    bEmpty = EmptyFolder(fld);
                    Console.WriteLine("\r\nFinished removing items from the CalRestore folder.");
                }
            }
            fldCalRestore = fld;
        }

        public static bool EmptyFolder(Folder fld)
        {
            int iOffset = 0;
            int iPageSize = 500;
            bool bMore = true;
            ItemView itemView = new ItemView(iPageSize, iOffset, OffsetBasePoint.Beginning);
            FindItemsResults<Item> results = null;
            char[] cSpin = new char[] { '/', '-', '\\', '|' };
            bool bRet = true;

            while (bMore)
            {
                int i = 0;
                int n = 0;
                results = fld.FindItems(itemView);
                foreach (Item item in results.Items)
                {
                    i++;
                    if (i % 5 == 0)
                    {
                        Console.SetCursorPosition(0, Console.CursorTop);
                        Console.Write("");
                        Console.Write(cSpin[n % 4]);
                        n++;
                    }
                    item.Delete(DeleteMode.HardDelete);
                }
                bMore = results.MoreAvailable;
                if (bMore)
                {
                    itemView.Offset += iPageSize;
                }
            }
            
            return bRet;
        }

        public static Folder FindFolder(string strFolder)
        {
            Folder fldSearch = null;
            FindFoldersResults fFRes = null;
            int iPageSize = 100;
            int iOffset = 0;
            bool bMore = true;

            FolderView view = new FolderView(iPageSize, iOffset, OffsetBasePoint.Beginning);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Shallow;   // should be a direct child so Shallow should get it

            // go get the results and find our folder
            while (bMore)
            {
                fFRes = exService.FindFolders(WellKnownFolderName.RecoverableItemsRoot, view);
                foreach (Folder fld in fFRes)
                {
                    if (fld.DisplayName == strFolder)
                    {
                        fldSearch = fld;
                        break;
                    }
                }
                // break out of the while loop if we got the folder
                if (fldSearch != null)
                {
                    break;
                }
                else
                {
                    bMore = fFRes.MoreAvailable;
                    if (bMore)
                    {
                        view.Offset += iPageSize;
                    }
                }
            }

            return fldSearch;
        }

        public static CalendarFolder FindCalFolder(string strFolder)
        {
            CalendarFolder fldSearch = null;
            FindFoldersResults fFRes = null;
            int iPageSize = 100;
            int iOffset = 0;
            bool bMore = true;

            FolderView view = new FolderView(iPageSize, iOffset, OffsetBasePoint.Beginning);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Shallow;   // should be a direct child so Shallow should get it

            // go get the results and find our folder
            while (bMore)
            {
                fFRes = exService.FindFolders(WellKnownFolderName.Calendar, view);
                foreach (CalendarFolder fld in fFRes)
                {
                    if (fld.DisplayName == strFolder)
                    {
                        fldSearch = fld;
                        break;
                    }
                }
                // break out of the while loop if we got the folder
                if (fldSearch != null)
                {
                    break;
                }
                else
                {
                    bMore = fFRes.MoreAvailable;
                    if (bMore)
                    {
                        view.Offset += iPageSize;
                    }
                }
            }

            return fldSearch;
        }

        // Go get an OAuth token to use Exchange Online 
        private static AuthenticationResult GetToken()
        {
            AuthenticationResult ar = null;
            AuthenticationContext ctx = new AuthenticationContext(strAuthCommon);

            try
            {
                ar = ctx.AcquireTokenAsync(strSrvURI, strClientID, new Uri(strRedirURI), new PlatformParameters(PromptBehavior.Always)).Result;
            }
            catch (Exception Ex)
            {
                var authEx = Ex.InnerException as AdalException;

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("An error occurred during authentication with the service:");
                Console.WriteLine(authEx.HResult.ToString("X"));
                Console.WriteLine(authEx.Message);
                Console.ResetColor();
            }
            return ar;
        }

        public static NameResolutionCollection DoResolveName(string strResolve)
        {
            NameResolutionCollection ncCol = null;
            try
            {
                ncCol = exService.ResolveName(strResolve, ResolveNameSearchLocation.DirectoryOnly, true);
            }
            catch (ServiceRequestException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error when attempting to resolve the name for " + strResolve + ":");
                Console.WriteLine(ex.Message);
                Console.ResetColor();
                return null;
            }

            return ncCol;
        }
    }
}
