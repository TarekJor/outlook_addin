using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;
using System.Runtime.Serialization;
using Microsoft.Office.Tools.Ribbon;
using WinForms = System.Windows.Forms;
using Res = Tabbles.OutlookAddIn.Properties.Resources;
using stdole;
using System.Drawing;
using System.Threading;
using System.Diagnostics;

namespace Tabbles.OutlookAddIn
{
    public delegate bool SendMessageToTabblesHandler(object message);

    public delegate bool IsAnyEmailSelectedHandler();

    class MenuManager
    {

        // SUJAYXML
      //   private XMLFileManager xmlFileManager;


        private const string CommandBarName = "Tabbles Toolbar";
        private const string ButtonIdTagUsingTabbles = "tagUsingTabbles";
        private const string ButtonIdOpenInTabbles = "openInTabbles";
        private const string ButtonIdTabblesSearch = "tabblesSearch";
        private const string ButtonIdSyncWithTabbles = "syncWithTabbles";
        private const string PropertyNameCategories = "Categories";
        private const string PropertyNameFlagRequest = "FlagRequest";

        public event SendMessageToTabblesHandler SendMessageToTabbles;
        public event System.Action StartSync;

        public readonly object syncObj = new object();

        private OutlookVersion outlookVersion;
        private string outlookPrefix;
        private CultureInfo outlookCulture;

        private Application outlookApp;
        private Explorers explorers;

        //keep the list members to avoid VSTO garbage collection problem
        private List<Explorer> explorerList;
        private List<CommandBarButton> buttonList;

        private List<MailItem> selectedMails;

        private Items currentFolderItems;

        //private ISet<string> onceItemChanged;

        public ISet<string> InternallyChangedMailIds
        {
            get;
            private set;
        }

        private bool trackItemMove = true;

        public TabblesRibbon Ribbon
        {
            set
            {
                value.TagUsingTabbles += (sender, args) =>
                {
                    TagUsingTabblesButtonClick();
                };
                value.OpenInTabbles += (sender, args) =>
                {
                    OpenInTabblesButtonClick();
                };
                value.TabblesSearch += (sender, args) =>
                {
                    TabblesSearch();
                };
                value.SyncWithTabbles += (sender, args) =>
                {
                    //RegistryManager.SetSyncPerformed(false);

                    // SUJAYXML
                    //xmlFileManager.SetSyncPerformed(false);

                    if (StartSync != null)
                    {
                        StartSync();
                    }
                };
                value.IsAnyEmailSelected += () =>
                {
                    return IsAnyEmailSelected(true);
                };
            }
        }

        public MenuManager(Application outlookApp)
        {
            // SUJAYXML
            //xmlFileManager = new XMLFileManager();

            this.outlookApp = outlookApp;
            this.outlookVersion = Utils.ParseMajorVersion(outlookApp);
            this.outlookPrefix = Utils.GetOutlookPrefix();

            this.explorerList = new List<Explorer>();
            this.buttonList = new List<CommandBarButton>();

            //this.onceItemChanged = new HashSet<string>();

            InternallyChangedMailIds = new HashSet<string>();

            //culture info for localization
            int languageId = outlookApp.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);
            this.outlookCulture = new CultureInfo(languageId);
            System.Threading.Thread.CurrentThread.CurrentUICulture = this.outlookCulture;

            CheckMenus();

            this.explorers = this.outlookApp.Explorers;
            this.explorers.NewExplorer += OnNewExplorer;

            FillItemsToListen();

            foreach (Explorer explorer in this.explorers)
            {
                AddExplorerListeners(explorer);
            }
        }

        #region Event handling
        private void OnNewExplorer(Explorer explorer)
        {
            AddExplorerListeners(explorer);
        }

        private void AddExplorerListeners(Explorer explorer)
        {
            this.explorerList.Add(explorer);
            explorer.BeforeItemPaste += TrackEmailMove;
            explorer.SelectionChange += UpdateSelectedEmails;

            explorer.FolderSwitch += () =>
                {
                    FillItemsToListen();
                };
        }

        private void TrackEmailMove(ref object clipboardContent, MAPIFolder target, ref bool cancel)
        {
            if (!this.trackItemMove) //prevent infinite loop
            {
                return;
            }

            if (clipboardContent is Selection)
            {
                List<MailItem> mails = new List<MailItem>();

                Selection selection = (Selection)clipboardContent;
                foreach (object itemObj in selection)
                {
                    if (itemObj is MailItem)
                    {
                        mails.Add((MailItem)itemObj);
                    }
                }

                if (mails.Count == 0)
                {
                    return;
                }

                bool movedFromStore = false;
                try
                {
                    foreach (MailItem mail in mails)
                    {
                        if (string.IsNullOrEmpty(mail.Categories))
                        {
                            continue;
                        }

                        if (mail.Parent is Folder)
                        {
                            Folder parent = (Folder)mail.Parent;
                            if (parent.StoreID != target.StoreID)
                            {
                                movedFromStore = true;
                                break;
                            }
                        }
                    }

                    if (!movedFromStore)
                    {
                        return;
                    }

                    if (!CheckTabblesRunning())
                    {
                        cancel = true;
                        WinForms.MessageBox.Show(Res.MsgTabblesIsNotRunning, Res.MsgCaptionTabblesAddIn);
                        return;
                    }

                    cancel = true;
                    this.trackItemMove = false;

                    foreach (MailItem mail in mails)
                    {
                        MailItem mailAfterMove = (MailItem)mail.Move(target);
                        Utils.ReleaseComObject(mailAfterMove);
                        //WinForms.MessageBox.Show(mail.EntryID + "\n\n" + mailAfterMove.EntryID);
                        //TODO: Maurizio: call Tabbles API at this point
                    }
                    this.trackItemMove = true;

                }
                finally
                {
                    foreach (MailItem mail in mails)
                    {
                        Utils.ReleaseComObject(mail);
                    }
                }
            }
        }

        private void UpdateSelectedEmails()
        {
            Selection selection = null;
            try
            {
                selection = this.outlookApp.ActiveExplorer().Selection;
                FillSelectedMails(selection);
            }
            catch (System.Exception)
            {
                //sometimes there could be an exception if it is something wrong with the folder
            }
        }

        public bool CheckTabblesRunning()
        {
            if (SendMessageToTabbles == null)
            {
                return false;
            }

            return SendMessageToTabbles(new INeedToPingTabbles());
        }

        public void SendEmailCategories(List<string> entryIds)
        {
            if (SendMessageToTabbles == null)
            {
                return;
            }

            foreach (string entryId in entryIds)
            {
                try
                {
                    MailItem mail = this.outlookApp.Session.GetItemFromID(entryId) as MailItem;
                    if (mail != null)
                    {
                        SendCategoriesToTabbles(mail);
                    }
                }
                catch (System.Exception)
                {
                }
            }
        }
        #endregion

        private void CheckMenus()
        {
            if (this.outlookVersion == OutlookVersion.OUTLOOK_2003 ||
                this.outlookVersion == OutlookVersion.OUTLOOK_2007)
            {
                CommandBar commandBar = null;
                try
                {
                    commandBar = this.outlookApp.ActiveExplorer().CommandBars[CommandBarName];
                    if (commandBar != null)
                    {
                        commandBar.Delete();
                    }
                }
                catch (System.Exception)
                {
                }

                commandBar = this.outlookApp.ActiveExplorer().CommandBars.Add(CommandBarName, MsoBarPosition.msoBarTop, Temporary: true);

                CommandBarButton tagUsingTabbles = CreateCommandBarButton(commandBar, Res.MenuTagUsingTabbles, ButtonIdTagUsingTabbles, "tag_using_tabbles");
                tagUsingTabbles.Click += tagUsingTabblesMenuButton_Click;
                this.buttonList.Add(tagUsingTabbles);

                CommandBarButton openEmailInTabbles = CreateCommandBarButton(commandBar, Res.MenuOpenInTabbles, ButtonIdOpenInTabbles, "open_in_tabbles");
                openEmailInTabbles.Click += openInTabblesMenuButton_Click;
                this.buttonList.Add(openEmailInTabbles);

                CommandBarButton tabblesSearch = CreateCommandBarButton(commandBar, Res.MenuTabblesSearch, ButtonIdTabblesSearch, "search");
                tabblesSearch.Click += tabblesSearch_Click;
                this.buttonList.Add(tabblesSearch);

                CommandBarButton syncWithTabbles = CreateCommandBarButton(commandBar, Res.MenuSyncWithTabbles, ButtonIdSyncWithTabbles, "sync_with_tabbles");
                syncWithTabbles.Click += syncWithTabbles_Click;
                this.buttonList.Add(syncWithTabbles);

                commandBar.Protection = MsoBarProtection.msoBarNoCustomize;
                commandBar.Visible = true;

                this.outlookApp.ItemContextMenuDisplay += outlookApp_ItemContextMenuDisplay;
            }
        }

        private CommandBarButton CreateCommandBarButton(CommandBar commandBar, string caption, string tag, string pictureAlias)
        {
            CommandBarButton button = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton);
            button.Caption = caption;
            button.Tag = tag;
            SetButtonPicture(button, pictureAlias + "_16_bmp", pictureAlias + "_16_mask");

            return button;
        }

        private void outlookApp_ItemContextMenuDisplay(CommandBar commandBar, Selection selection)
        {
            if (IsAnyEmailSelected(true))
            {
                CommandBarButton tagUsingTabbles = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
                tagUsingTabbles.Caption = Res.MenuTagUsingTabbles;
                tagUsingTabbles.Click += tagUsingTabblesContextMenu_Click;
                SetButtonPicture(tagUsingTabbles, "tag_using_tabbles_16_bmp", "tag_using_tabbles_16_mask");
                this.buttonList.Add(tagUsingTabbles);

                if (this.selectedMails != null && this.selectedMails.Count == 1)
                {
                    CommandBarButton openInTabbles = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
                    openInTabbles.Caption = Res.MenuOpenInTabbles;
                    openInTabbles.Click += openInTabblesContextMenu_Click;
                    SetButtonPicture(openInTabbles, "open_in_tabbles_16_bmp", "open_in_tabbles_16_mask");
                    this.buttonList.Add(openInTabbles);
                }
            }
        }

        private void SetButtonPicture(CommandBarButton button, string imageName, string maskName)
        {
            IPictureDisp picture = GetPictureDispFromResource(imageName);
            if (picture != null)
            {
                button.Style = MsoButtonStyle.msoButtonIconAndCaption;
                button.Picture = picture;
                IPictureDisp mask = GetPictureDispFromResource(maskName);
                if (mask != null)
                {
                    button.Mask = mask;
                }
            }
            else
            {
                button.Style = MsoButtonStyle.msoButtonCaption;
            }
        }

        private void tagUsingTabblesMenuButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            TagUsingTabblesButtonClick();
        }

        private void TagUsingTabblesButtonClick()
        {
            if (IsAnyEmailSelected(true))
            {
                TagUsingTabbles(this.selectedMails);
            }
        }

        private void tagUsingTabblesContextMenu_Click(CommandBarButton ctrl, ref bool cancelDefault)
        {
            if (IsAnyEmailSelected(false))
            {
                TagUsingTabbles(this.selectedMails);
            }
        }

        private void TagUsingTabbles(List<MailItem> mails)
        {
            if (SendMessageToTabbles == null)
            {
                return;
            }

            var emails = (from MailItem mi in this.selectedMails
                          select new Generic
                          {
                              name = mi.Subject,
                              commandLine = this.outlookPrefix + mi.EntryID,
                              icon = new IconOther(),
                              showCommandLine = false
                          }).ToList();

            SendMessageToTabbles(new INeedToTagGenericsWithTabblesQuickTagDialog()
            {
                gens = emails
            });
        }

        private void openInTabblesMenuButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OpenInTabblesButtonClick();
        }

        private void OpenInTabblesButtonClick()
        {
            if (IsAnyEmailSelected(true))
            {
                OpenInTabbles(this.selectedMails[0]);
            }
        }

        private void openInTabblesContextMenu_Click(CommandBarButton ctrl, ref bool cancelDefault)
        {
            if (IsAnyEmailSelected(false))
            {
                OpenInTabbles(this.selectedMails[0]);
            }
        }

        private void OpenInTabbles(MailItem mail)
        {
            if (SendMessageToTabbles == null)
            {
                return;
            }

            var email = new Generic
            {
                name = mail.Subject,
                commandLine = this.outlookPrefix + mail.EntryID,
                icon = new IconOther(),
                showCommandLine = false
            };
            SendMessageToTabbles(new INeedToOpenGenericInTabbles()
            {
                gen = email
            });
        }

        private void tabblesSearch_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            TabblesSearch();
        }

        private void TabblesSearch()
        {
            if (SendMessageToTabbles == null)
            {
                return;
            }

            SendMessageToTabbles(new INeedToOpenSearch());
        }

        private void syncWithTabbles_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (StartSync != null)
            {
                StartSync();
            }
        }

        private bool IsAnyEmailSelected(bool fillAtFirst)
        {
            if (fillAtFirst)
            {
                try
                {
                    FillSelectedMails(this.outlookApp.ActiveExplorer().Selection);
                }
                catch
                {
                    return false;
                }
            }

            return (this.selectedMails != null && this.selectedMails.Count > 0);
        }

        private void FillSelectedMails(Selection selection)
        {
            if (selection.Count > 0 && selection[1] is MailItem)
            {
                if (this.selectedMails == null)
                {
                    this.selectedMails = new List<MailItem>();
                }
                else
                {
                    this.selectedMails.Clear();
                }

                foreach (var sel in selection)
                {
                    if (sel is MailItem)
                    {
                        MailItem mail = (MailItem)sel;
                        this.selectedMails.Add(mail);
                    }
                }
            }
            else if (this.selectedMails != null)
            {
                this.selectedMails.Clear();
            }
        }

        private void FillItemsToListen()
        {
            if (this.currentFolderItems != null)
            {
                try
                {
                    this.currentFolderItems.ItemChange -= Items_ItemChange;
                }
                catch (System.Exception)
                {
                }
            }

            Folder currentFolder = (Folder)this.outlookApp.ActiveExplorer().CurrentFolder;

            if (currentFolder != null)
            {
                this.currentFolderItems = currentFolder.Items;

                //avoid double adding
                this.currentFolderItems.ItemChange -= Items_ItemChange;
                this.currentFolderItems.ItemChange += Items_ItemChange;
            }
        }

        private void Items_ItemChange(object item)
        {
            if (item is MailItem)
            {
                MailItem mail = (MailItem)item;
                string mailId = mail.EntryID;
                lock (this.syncObj)
                {
                    //if (this.onceItemChanged.Contains(mailId))
                    //{
                    //    this.onceItemChanged.Remove(mailId);
                    //}
                    //else
                    if (InternallyChangedMailIds.Contains(mailId))
                    {
                        InternallyChangedMailIds.Remove(mailId);
                        //this.onceItemChanged.Add(mailId);
                    }
                    else
                    {
                        if (SendCategoriesToTabbles(mail))
                        {
                            //this.onceItemChanged.Add(mailId);
                        }
                    }
                }
            }
        }

        private bool SendCategoriesToTabbles(MailItem mail)
        {
            if (SendMessageToTabbles == null)
            {
                return false;
            }

            var categoriesWithColors = new Dictionary<string, string>();
            string[] categories = Utils.GetCategories(mail);
            foreach (string categoryName in categories)
            {
                try
                {
                    Category category = this.outlookApp.Session.Categories[categoryName];
                    if (category != null)
                    {
                        string categoryRgb = Utils.GetRgbFromOutlookColor(category.Color);
                        categoriesWithColors[categoryName] = categoryRgb;
                    }
                }
                catch (System.Exception)
                {
                    //ignore the category
                }
            }

            if (!string.IsNullOrEmpty(mail.FlagRequest))
            {
                categoriesWithColors[mail.FlagRequest] = Utils.GetRgbForFlagRequest(mail.FlagRequest);
            }

            var message = new GenericChangedSomeCategory
            {
                categoriesWithColors = categoriesWithColors,
                gen = new Generic
                {
                    name = mail.Subject,
                    commandLine = this.outlookPrefix + mail.EntryID,
                    icon = new IconOther(),
                    showCommandLine = false
                }
            };

            return SendMessageToTabbles(message);
        }

        /// <summary>
        /// Adds Entry ID of MailItem to skip listening its changes, for instance if changes are done programmatically.
        /// </summary>
        /// <param name="entryId"></param>
        //public void AddEntryIdToSkip(string entryId)
        //{
        //    this.itemsToSkipChanges.Add(entryId);
        //}

        private IPictureDisp GetPictureDispFromResource(string resourceName)
        {
            object resource = Res.ResourceManager.GetObject(resourceName);
            if (resource is Image)
            {
                return ImageConverter.GetPictureDisp((Image)resource);
            }

            return null;
        }
    }
}
