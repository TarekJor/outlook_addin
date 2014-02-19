﻿using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Res = Tabbles.OutlookAddIn.Properties.Resources;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Media;
using Outlook = Microsoft.Office.Interop.Outlook;

using System.Xml.Linq;

#region Sujay
//using Redemption;

#endregion

namespace Tabbles.OutlookAddIn
{
    public partial class ThisAddIn
    {
        private static readonly string[] OutlookCmdSeparator = new string[] { @"/select outlook:" };

        private const string SearchResultsFolderName = "Tabbles search results";

        private MenuManager menuManager;
        private FolderManager folderManager;
        private ItemManager itemManager;
        private SyncManager syncManager;
        private TabblesRibbon ribbon;

        //SUJAYXML
        //  private XMLFileManager xmlFileManager;

        #region Sujay
        //private RDOSession rdoSession; 
        #endregion

        private BinaryFormatter formatter = new BinaryFormatter();
        private Thread listenerThread;

        Outlook.Application mApplication;
        Outlook.NameSpace mSession;
        Outlook.Folders mFolders;
        

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                //string redemptionDllPath = @"D:\Projects\Tabbles\TabblesOutlookAddIn\TabblesLibrary\";
                //RedemptionLoader.DllLocation32Bit = redemptionDllPath + "Redemption.dll";
                //RedemptionLoader.DllLocation64Bit = redemptionDllPath + "Redemption64.dll";

                Logger.Log("Outlook plugin initialized.");


                //SUJAYXML
                // xmlFileManager = new XMLFileManager();

                // SUJAYXML
                //xmlFileManager.CreateSettingsFile();

                mApplication = this.Application;

                this.menuManager = new MenuManager(mApplication);

                this.menuManager.Ribbon = this.ribbon;
                ribbon.mMenuManager = menuManager;

                this.itemManager = new ItemManager();


                mSession = mApplication.Session;

                mFolders = mSession.Folders;

                this.syncManager = new SyncManager(mFolders);
                this.syncManager.SendEmailCategories += this.menuManager.SendEmailCategories;

                //this.menuManager.StartSync += delegate
                //{
                //    StartSyncThread();
                //};

                this.listenerThread = new Thread(ListenTabblesEvents);
                this.listenerThread.Start();

                //Application.AdvancedSearchComplete += new ApplicationEvents_11_AdvancedSearchCompleteEventHandler(Application_AdvancedSearchComplete);

                //  if (!RegistryManager.IsSyncPerformed() && !RegistryManager.IsDontAskForSync())

                //StartSyncThread();

                // SUJAYXML

                //     if (!xmlFileManager.IsSyncPerformed() && !xmlFileManager.IsDontAskForSync())
                //if (!RegistryManager.IsSyncPerformed() && !RegistryManager.IsDontAskForSync())
                //{
                //    StartSyncThread();
                //}


                //var ds = Application.Session.DefaultStore;
                foreach (Folder f in mFolders)
                {
                    var it = f.Items;
                    mItems.Add( it); // fissare questo è fondamentale! se uso f.Items ogni volta non funziona. http://www.outlookbanter.com/outlook-vba/7399-mapi-folder-items-itemchange-event.html
                    //var name = f.Name;
                    //mItems.Add( items); // prevent garbage collection
                    
                    it.ItemChange += Items_ItemChange;

                    it.ItemAdd += mItems_ItemAdd;
                    it.ItemRemove += mItems_ItemRemove;

                    
                    //foreach (var x in f.Items)
                    //{
                    //    if (x is MailItem)
                    //    {
                    //        var x2 = (MailItem)x;
                    //        mMails.Add(x2);

                    //        x2.PropertyChange += x2_PropertyChange;

                    //    }

                    //}
                }
            }
            catch (System.Exception ex)
            {
                Logger.Log(ex.ToString());
            }
        }

        void mItems_ItemRemove()
        {
            var y = 5;
        }

        void mItems_ItemAdd(object Item)
        {
            var y = 5;
        }

        void x2_PropertyChange(string Name)
        {
            var y = 3;
        }

        //static List<Folder> mFolders = new HashSet<Folder>(); // prevents garbage collection. otherwise itemchange is not fired.

        List<Items> mItems = new List<Items>();
        MAPIFolder mFolder;
        HashSet<MailItem> mMails = new HashSet<MailItem>(); // prevents garbage collection. otherwise itemchange is not fired.

        void Items_ItemChange(object Item)
        {
            var y = 3;
        }

        private void StartSyncThread()
        {
            System.Action syncAction = this.syncManager.GetSyncAction();
            syncAction();
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            this.ribbon = new TabblesRibbon();



            return this.ribbon;
        }

        //private bool OnSendMessageToTabbles(object message)
        //{
        //    return SendMessageToTabblesBlocking(message);
        //}

        //private bool SendMessageToTabblesBlocking(object msg, bool retry = false)
        //{
        //    try
        //    {
        //        // I commented this block because this function should should never fail without showing an error message box.
        //        //if (msg.GetType().GetCustomAttributes(typeof(SerializableAttribute), false).Length == 0)
        //        //{
        //        //    return false;
        //        //}

        //        if (this.outlookToTabblesClientPipe == null || retry)
        //        {
        //            this.outlookToTabblesClientPipe = new NamedPipeClientStream(".", "OutlookToTabblesPipe",
        //                PipeDirection.Out, PipeOptions.Asynchronous);
        //            Logger.Log("connecting to Tabbles pipe server...");
        //            this.outlookToTabblesClientPipe.Connect(200); // blocks the thread
        //            Logger.Log("connected.");
        //        }

        //        Logger.Log("sendMessageToTabblesBlocking: serialize: " + msg.GetType().ToString());
        //        this.formatter.Serialize(this.outlookToTabblesClientPipe, msg);
        //        this.outlookToTabblesClientPipe.Flush();

        //        return true;
        //        //logFile.Print("sendMessageToTabblesBlocking: sent");
        //    }
        //    catch (TimeoutException)
        //    {
        //        string str = "Tabbles plugin not active. Cannot send message to Tabbles: " + msg.GetType().ToString();
        //        Logger.Log(str);

        //        try
        //        {
        //            this.outlookToTabblesClientPipe.Dispose();
        //        }
        //        catch (System.Exception)
        //        { }
        //        finally
        //        {
        //            this.outlookToTabblesClientPipe = null;
        //        }

        //        return false;
        //    }
        //    catch (System.Exception)
        //    {
        //        if (!retry)
        //        {
        //            try
        //            {
        //                this.outlookToTabblesClientPipe.Dispose();
        //            }
        //            catch (System.Exception)
        //            { }
        //            finally
        //            {
        //                this.outlookToTabblesClientPipe = null;
        //            }

        //            //try once more to re-connect the pipe
        //            if (SendMessageToTabblesBlocking(msg, true))
        //            {
        //                return true;
        //            }
        //            else
        //            {
        //                Logger.Log("The Tabbles plugin for Outlook is not running.");
        //            }
        //        }

        //        return false;
        //    }
        //}

        private void handleMessageFromTabbles(XDocument xdoc)
        {
            var root = xdoc.Root;
            if (root.Name.LocalName == "emails_tagged")
            {
                var emails = root.Elements("email");
                var tags = root.Elements("tag");

                foreach (var email in emails)
                {
                    var cmdLine = email.Attribute("command_line").Value;
                    // I have to tag the same email with categories corresponding to the tags
                    string[] arguments = cmdLine.Split(OutlookCmdSeparator, StringSplitOptions.None);

                    string entryId = arguments[1];

                    MailItem mail = (MailItem)Application.Session.GetItemFromID(entryId);

                    string[] currentCategories;
                    if (mail.Categories != null)
                    {
                        currentCategories = Utils.GetCategories(mail);
                    }
                    else
                    {
                        currentCategories = new string[0];
                    }

                    var tagsToAddWithColors = (from tag in tags
                                               let tagName = tag.Attribute("name").Value
                                               let tagColor = tag.Attribute("color").Value
                                               where currentCategories.All(cat => cat != tagName)
                                               select new { name = tagName, color = tagColor }).ToList();

                    if (!tagsToAddWithColors.Any())
                    {
                        continue;
                    }

                    foreach (var tag in tagsToAddWithColors)
                    {
                        Category cat;
                        if (!CategoryExists(tag.name))
                        {
                            cat = this.Application.Session.Categories.Add(tag.name);
                        }
                        else
                        {
                            cat = this.Application.Session.Categories[tag.name];
                        }

                        //change colors for all categories, in case if they were changed in Tabbles
                        cat.Color = Utils.GetOutlookColorFromRgb(tag.color);
                    }

                    var tagsToAdd = (from x in tagsToAddWithColors
                                     select x.name);
                    IEnumerable<string> newCats = tagsToAdd.Concat<string>(currentCategories);
                    // todo newcats is empty: ???? check, are they
                    mail.Categories = newCats.Aggregate((a, b) => a + ";" + b);

                    this.menuManager.InternallyChangedMailIds.Add(entryId);

                    mail.Save();

                }

            }
            else if (root.Name.LocalName == "emails_untagged")
            {
                var emails = root.Elements("email");
                var tags = root.Elements("tag");
                foreach (var email in emails)
                {
                    var cmdLine = email.Attribute("command_line").Value;

                    // I have to tag the same email with categories corresponding to the tags
                    string[] arguments = cmdLine.Split(OutlookCmdSeparator, StringSplitOptions.None);

                    string entryId = arguments[1];

                    MailItem mail = (MailItem)Application.Session.GetItemFromID(entryId);

                    string[] currentCategories;
                    if (mail.Categories != null)
                    {
                        currentCategories = Utils.GetCategories(mail);
                    }
                    else
                    {
                        continue;
                    }


                    var tagnames = (from tag in tags
                                    select tag.Attribute("name").Value);
                    var newCats = currentCategories.Except<string>(tagnames).ToList();

                    if (newCats.Any())
                    {
                        mail.Categories = newCats.Aggregate((a, b) => a + ";" + b);
                    }
                    else
                    {

                        mail.Categories = null;
                    }


                    this.menuManager.InternallyChangedMailIds.Add(entryId);

                    mail.Save();
                }

            }
            else if (root.Name.LocalName == "find_emails_which_have_these_tags")
            {
                var tags = root.Elements("tag");

                var tagnames = (from tag in tags
                                select tag.Attribute("name").Value);

                //MsgOpenMailsWithTags msgOpenMailsWithTags = (MsgOpenMailsWithTags)messageObj;
                //if (msgOpenMailsWithTags.tags != null)
                //{
                SearchByCategories(tagnames);
                //}
            }
            //else if (root.Name.LocalName == "tag_created")
            //{
            //    //MsgAtomKeyCreated msgAtomKeyCreated = (MsgAtomKeyCreated)messageObj;
            //    //string categoryName = msgAtomKeyCreated.AtomKeyName;

            //    //Category category;
            //    //if (!CategoryExists(categoryName))
            //    //{
            //    //    category = this.Application.Session.Categories.Add(categoryName);
            //    //}
            //    //else
            //    //{
            //    //    category = this.Application.Session.Categories[categoryName];
            //    //}

            //    //category.Color = Utils.GetOutlookColorFromRgb(msgAtomKeyCreated.AtomKeyColor);

            //    //Logger.Log("detected ak created: " + msgAtomKeyCreated.AtomKeyName);
            //}
            else if (root.Name.LocalName == "tag_deleted")
            {
                Logger.Log("detected ak deleted");
            }
            else
            {
                Logger.Log("message from Tabbles not recognized: " + root.ToString());
            }
        }


        private void ListenTabblesEvents()
        {

            while (true)
            {
                try
                {
                    using (var pipeServer = new NamedPipeServerStream("TABBLES_PIPE_TO_OUTLOOK", PipeDirection.InOut)) // inout per prevenire il bug che succedeva nell'altro verso. cioè, con solo in, dà unauthorizedaccessexception.
                    {
                        Logger.Log("Waiting for Tabbles to connect to outlook pipe...");
                        pipeServer.WaitForConnection(); //blocking

                        Logger.Log("Connection established.");

                        var xdoc = XDocument.Load(pipeServer);

                        handleMessageFromTabbles(xdoc);
                    }


                }
                catch (System.Exception e)
                {
                    Logger.Log("exception - restarting pipe server");
                }
            }

        }

        private bool CategoryExists(string categoryName)
        {
            try
            {
                Category category =
                    this.Application.Session.Categories[categoryName];

                return category != null;
            }
            catch
            {
                return false;
            }
        }

        public void SearchByCategories(IEnumerable<string> categories)
        {

            Outlook.Explorer explorer = Application.Explorers.Add(
                   Application.Session.GetDefaultFolder(
                   Outlook.OlDefaultFolders.olFolderInbox)
                   as Outlook.Folder,
                   Outlook.OlFolderDisplayMode.olFolderDisplayNormal);

            var cats = (from c in categories
                        select "category:\"" + c + "\"").Aggregate( (a, b) => a + " AND " + b);
            
            explorer.Search(cats, Outlook.OlSearchScope.olSearchScopeAllFolders);
            explorer.Display();


            //Folder currentFolder = (Folder)Application.ActiveExplorer().CurrentFolder;



            //Folder rootFolder;


            //if (currentFolder != null)
            //{
            //    rootFolder = (Folder)currentFolder.Store.GetRootFolder();
            //}
            //else
            //{
            //    rootFolder = (Folder)Application.Session.Folders[1];
            //}




            ////example: ("urn:schemas-microsoft-com:office:office#Keywords" = 'aa' OR "urn:schemas-microsoft-com:office:office#Keywords" = 'bb')
            //int count = categories.Count<string>();
            //StringBuilder filterSql = new StringBuilder("(");
            //if (count > 0)
            //{
            //    filterSql.AppendFormat("\"urn:schemas-microsoft-com:office:office#Keywords\" = '{0}'", categories.First<string>());
            //}
            //else
            //{
            //    return;
            //}

            //for (int i = 1; i < count; i++)
            //{
            //    // andrej aveva messo OR. a me sembra senza senso.
            //    filterSql.Append(" AND ").AppendFormat("\"urn:schemas-microsoft-com:office:office#Keywords\" = '{0}'", categories.ElementAt<string>(i));
            //}
            //filterSql.Append(")");

            //#region old comment by andrej
            ////-- We use Redemption instead of these code (together with AdvancedSearchComplete event, see another Commented out section)
            ////-- Currently there is a problem with calling Results.Save() for search on a folder of non-default store
            ////See: http://social.msdn.microsoft.com/Forums/en-US/outlookdev/thread/7d1d3494-988f-4c42-a391-e732b5dfb2c6

            ////string folderStr = string.Format("'{0}'", rootFolder.FolderPath);

            ////string logMessage = string.Format("Started search with filter {0} in folder {1} ...", filter.ToString(), folderStr);
            ////this.logger.Log(logMessage);

            ////Application.AdvancedSearch(folderStr, filter.ToString(), true, "Tabbles categories");
            ////--------------------------------------------------------------------------------------
            //#endregion

            //var performSearch = new System.Action(() =>
            //    {
            //            #region Sujay
            //            //if (this.rdoSession == null)
            //            //{
            //            //    this.rdoSession = RedemptionLoader.new_RDOSession();
            //            //}
            //            //if (!this.rdoSession.LoggedOn)
            //            //{
            //            //    this.rdoSession.Logon();
            //            //}

            //            //RDOStore2 store = (RDOStore2)this.rdoSession.GetStoreFromID(rootFolder.StoreID);

            //            //oInbox = oApp.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);


            //            //NameSpace olNS = this.Application.GetNamespace("MAPI");
            //            //Store olStore = olNS.GetStoreFromID(rootFolder.StoreID);

            //            //MAPIFolder olSearchFolder;

            //            //  olStore.

            //            //  Application.AdvancedSearchComplete -= new ApplicationEvents_11_AdvancedSearchCompleteEventHandler(Application_AdvancedSearchComplete);

            //            string folderStr = string.Format("'{0}'", rootFolder.FolderPath);
            //            var olSearch = Application.AdvancedSearch(Scope: folderStr, Filter: filterSql.ToString(), SearchSubFolders: true, Tag: SearchResultsFolderName);
            //            olSearch.Save(SearchResultsFolderName);

            //            //     olSearchFolder = olSearch.Save("Sujay Search");

            //            //Application.AdvancedSearchComplete -= new ApplicationEvents_11_AdvancedSearchCompleteEventHandler(Application_AdvancedSearchComplete);

            //            //store.OnSearchComplete += store_OnSearchComplete;



            //            //MAPIFolder olFolderFromID = olNS.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID);


            //            //RDOFolder folder = this.rdoSession.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID);

            //            // Sujay code

            //            //store.Searches.AddCustom(SearchResultsFolderName, filterSql.ToString(), folder, true); 


            //            #endregion

            //    });

            //Folders searchFolders = rootFolder.Store.GetSearchFolders();
            ////if (searchFolders != null)
            ////{
            ////    if (this.folderManager == null)
            ////    {
            ////        this.folderManager = new FolderManager();
            ////    }

            ////    //in case if there is a search folder
            ////    this.folderManager.RemoveFolderByName(folders: searchFolders, name: SearchResultsFolderName, callback: performSearch);
            ////}
            ////else
            ////{
            //    //in case if there is no any search folder
            //    performSearch();
            ////}

            return;
        }

        //private void store_OnSearchComplete(string searchFolderID)
        //{
        //    #region Sujay
        //    //Folder searchFolder = (Folder)Application.Session.GetFolderFromID(searchFolderID);
        //    //if (this.rdoSession != null && this.rdoSession.LoggedOn)
        //    //{
        //    //    RDOStore2 store = (RDOStore2)this.rdoSession.GetStoreFromID(searchFolder.StoreID);
        //    //    store.OnSearchComplete -= store_OnSearchComplete;
        //    //}

        //    //Application.ActiveExplorer().CurrentFolder = searchFolder; 
        //    #endregion
        //}

        #region Commented out
        //see comment in SearchByCategories() for the explanation

        //private void Application_AdvancedSearchComplete(Search search)
        //{
        //    #region Sujay Comments
        //    //string logMessage = string.Format("Search is completed with {0} results.", search.Results.Count.ToString());
        //    ////this.logger.Log(logMessage);

        //    //if (search.Results.Count != 0)
        //    //{
        //    //    search.Save("Sujay Search");
        //    //    return;
        //    //}

        //    if (search.Results.Count == 0)
        //    {
        //        MessageBox.Show(Res.MsgNoResultsFound);
        //    }
        //    else
        //    {
        //        //search.Save("Sujay Search");

        //        Folders searchFolders = null;
        //        MailItem aMail = search.Results[1] as MailItem;
        //        Folder aFolder = aMail.Parent as Folder;
        //        searchFolders = aFolder.Store.GetSearchFolders();

        //        var showResultsAction = new System.Action(() =>
        //            {
        //                Folder searchResultsFolder = (Folder)search.Save(SearchResultsFolderName);
        //                Application.ActiveExplorer().CurrentFolder = searchResultsFolder;
        //            });

        //        //if (searchFolders != null)
        //        //{
        //        //    if (this.folderManager == null)
        //        //    {
        //        //        this.folderManager = new FolderManager();
        //        //    }

        //        //    //in case if there is a search folder
        //        //    this.folderManager.RemoveFolderByName(searchFolders, SearchResultsFolderName, showResultsAction);
        //        //}
        //        //else
        //        //{
        //        //in case if there is no any search folder
        //        showResultsAction();
        //        //}

        //        return;

        //        //give some response in any case
        //        //MessageBox.Show(Res.MsgNoResultsFound);
        //    }
        //    #endregion

        //    //MessageBox.Show(" In advanced search");

        //    //Application.AdvancedSearchComplete -= new ApplicationEvents_11_AdvancedSearchCompleteEventHandler(Application_AdvancedSearchComplete);

        //    //  Application.ActiveExplorer().CurrentView = searchFolder.Application.ActiveExplorer().CurrentView;//  = searchFolder.f;
        //}
        #endregion

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

            //Application.AdvancedSearchComplete -= new ApplicationEvents_11_AdvancedSearchCompleteEventHandler(Application_AdvancedSearchComplete);
            Logger.Dispose();
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
