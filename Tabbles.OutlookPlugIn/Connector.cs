using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO.Pipes;
using System.Runtime.Serialization.Formatters.Binary;
using Tabbles.OutlookAddIn.Common.Messages;
using Microsoft.FSharp.Collections;
using System.Runtime.Serialization;
using Microsoft.FSharp.Core;

namespace Tabbles.OutlookPlugIn
{



    public class Connector : TabblesApi.IInitializable, TabblesApi.IEventListener, TabblesApi.IPlugin
    {
        private BinaryFormatter formatter = new BinaryFormatter();
        private NamedPipeClientStream tabblesToOutlookClientPipe;

        private X compute<X>(Func<X> ac)
        {
            return ac.Invoke();
        }



        public void Initialize()
        {
            myLog.printDisp("outlook system: tabbles plugin: initialize");
            Thread listenerThread = new Thread(ListenOutlookEvents);
            listenerThread.Start();
        }

        private void SendMessageToOutlookBlocking(object message, bool retry = false)
        {
            try
            {
                myLog.printDisp("outlook system: tabbles plugin: sending message to outlook: " + message.GetType().ToString());
                if (this.tabblesToOutlookClientPipe == null || retry)
                {
                    this.tabblesToOutlookClientPipe = new NamedPipeClientStream(".", "TabblesToOutlookPipe", PipeDirection.Out,
                        PipeOptions.Asynchronous);
                    this.tabblesToOutlookClientPipe.Connect(1500); // blocks the thread
                }

                this.formatter.Serialize(this.tabblesToOutlookClientPipe, message);
                this.tabblesToOutlookClientPipe.Flush();
            }
            catch (TimeoutException)
            {
                return;
            }
            catch (Exception)
            {
                if (!retry)
                {
                    try
                    {
                        this.tabblesToOutlookClientPipe.Dispose();
                    }
                    catch (Exception)
                    {
                    }

                    //try once more
                    SendMessageToOutlookBlocking(message, true);
                }
            }
        }

        private void ListenOutlookEvents()
        {
            myLog.printDisp("outlook system: tabbles plugin: in server thread.");

            NamedPipeServerStream outlookToTabblesServerPipe;
            while (true)
            {
                // I need to set up the pipe to talk to the outlook plugin.
                outlookToTabblesServerPipe = new NamedPipeServerStream("OutlookToTabblesPipe", PipeDirection.In);
                myLog.printDisp("outlook system: tabbles plugin: server up, waiting for connection to the pipe...");
                outlookToTabblesServerPipe.WaitForConnection(); // blocking

                myLog.printDisp("outlook system: tabbles plugin: a client connected.");

                while (true)
                {
                    object messageObj = null;

                    try
                    {
                        messageObj = this.formatter.Deserialize(outlookToTabblesServerPipe);
                    }
                    catch (SerializationException ex)
                    {
                        outlookToTabblesServerPipe.Dispose(); // otherwise I will get error at the next call to new NamedPipeServerStream("PipeOutlookToTabbles", PipeDirection.In);

                        string errMessage = ex.GetType().ToString() + " --- " + ex.Message;
                        myLog.printDisp(errMessage);
                        //TabblesApi.API.ExecuteInGuiThread(
                        //    (Action)(() => { popup.showPopupNoLists("Error", "deserializationFailed: " + errMessage, Tabbles_decl.popupKind.PkNormal, popup.color.CError); }));

                        break; //create the pipe server again
                    }

                    // the outlook addin sent some event which we have to process.
                    //myLog.printDisp("outlook system: tabbles plugin: message received...");

                    //if (messageObj is INeedToTagGenericWithKnownTags)
                    //{
                    //    myLog.printDisp("outlook system: tabbles plugin: message was: i need to tag generic with known tags...");
                    //    var x = (INeedToTagGenericWithKnownTags)messageObj;
                    //    // tag the same email in tabbles
                    //    var cts = new List<Tabbles_decl.ct_key2>();
                    //    var ct2 = Tabbles_decl.ct_key2.NewCt2Gen(new Tabbles_decl.generic_key2(x.gen.commandLine, x.gen.name));
                    //    cts.Add(ct2);

                    //    var ctsl = SeqModule.ToList(cts);
                    //    var tags = SeqModule.ToList(x.tags);
                    //    var ret = db.tagCtsNonRecursive(ctsl, tags, db.wakeUpThreads.DoNotWakeUpThreads);

                    //}
                    if (messageObj is INeedToTagGenericsWithTabblesQuickTagDialog)
                    {
                        myLog.printDisp("outlook system: tabbles plugin: message was: i need to tag generic with tabbles quick dialog...");
                        var x = (INeedToTagGenericsWithTabblesQuickTagDialog)messageObj;

                        var cts2 = (from g in x.gens
                                    select Tabbles_decl.ct_key2.NewCt2Gen(new Tabbles_decl.generic_key2(g.commandLine, g.name)));

                        TabblesApi.API.ExecuteActionInGuiThread(() =>
                            {
                                Tabbles_logic.uiQuickTagOrUntagCts(SeqModule.ToList(cts2), Misc.whoCalledQuickCategorize.WcqPopupMessage, false);
                            });
                    }
                    else if (messageObj is INeedToOpenGenericInTabbles)
                    {
                        myLog.printDisp("outlook system: tabbles plugin: message was: i need to open generic in tabbles...");
                        var x = (INeedToOpenGenericInTabbles)messageObj;

                        var ct = Tabbles_decl.ct_key.CtGen.NewCtGen(x.gen.commandLine);

                        if (db.ctExists(db.mergedLocal.LocalDb, ct))
                        {
                            var cd = db.getDataOfCt(db.mergedLocal.LocalDb, ct);
                            var aks = Tabbles_logic.aksOfCt(ct, cd);
                            if (aks.IsEmpty)
                            {
                                TabblesApi.API.ExecuteActionInGuiThread(() =>
                                {
                                    popup.showPopupNoLists("Cannot open email", "This email is not yet categorized in Tabbles.", Tabbles_decl.popupKind.PkNormal, popup.color.CError);
                                });
                            }
                            else
                            {
                                TabblesApi.API.ExecuteActionInGuiThread(() =>
                                {
                                    Tabbles_logic.openTabblesWindowWithGivenAks(aks);
                                });
                            }
                        }
                        else
                        {
                            TabblesApi.API.ExecuteActionInGuiThread(() =>
                            {
                                popup.showPopupNoLists("Cannot open email", "This email is not yet categorized in Tabbles.", Tabbles_decl.popupKind.PkNormal, popup.color.CError);
                            });
                        }

                    }
                    else if (messageObj is GenericChangedSomeCategory)
                    {
                        myLog.printDisp("outlook system: tabbles plugin: message was: mail item changed some category...");
                        var x = (GenericChangedSomeCategory)messageObj;

                        myLog.printDisp("Generic: " + x.gen.commandLine + "\nCategories are the following: ");
                        foreach (var catWithColor in x.categoriesWithColors)
                        {
                            myLog.printDisp(catWithColor.ToString());
                        }

                        TabblesApi.API.LockDbA(() =>
                        {
                            var ct = Tabbles_decl.ct_key.NewCtGen(x.gen.commandLine);
                            var newTags = x.categoriesWithColors.Keys;


                            // TODO maurizio: debug. still does not work

                            // 1. create in tabbles'db all the tabbles that are needed
                            {
                                foreach (var catName in newTags)
                                {
                                    //myLog.printDisp("debug_outlook_plugin: " + catName);
                                    var fk = Tabbles_decl.formula_key.FkAtomic.NewFkAtomic
                                                (Tabbles_decl.atom_key.NewAsk(Tabbles_decl.atom_set_key.NewAskLabel(catName)));
                                    if (!db.formulaExists(db.mergedLocal.LocalDb, fk))
                                    {
                                        //myLog.printDisp("debug_outlook_plugin: formula does not exist");
                                        var colorId = compute<int>(() =>
                                            {
                                                var colStr = x.categoriesWithColors[catName];
                                                var colorKeysInTabbles = db.getColors(db.mergedLocal.LocalDb);
                                                var colIdsWithThatColor = (from co in colorKeysInTabbles
                                                                           let cad = db.getDataOfColor(db.mergedLocal.LocalDb, co)
                                                                           where cad.cad_cat_color == colStr
                                                                           select co).ToList();

                                                int colorID;
                                                if (colIdsWithThatColor.Count == 0)
                                                {
                                                    // There is no color with that color. Create a new color
                                                    var indexOfNewColor = Microsoft.FSharp.Collections.ListModule.Max(
                                                        db.getColors(db.mergedLocal.LocalDb)) + 1;
                                                    var cad = new Tabbles_decl.category_data(colStr, FSharpOption<string>.None);
                                                    //myLog.printDisp("debug_outlook_plugin: creating color ");
                                                    db.setDataOfColor(indexOfNewColor, cad);
                                                    return indexOfNewColor;

                                                }
                                                else
                                                {
                                                    //myLog.printDisp("debug_outlook_plugin: color exists");
                                                    // there is already a color with that color. use that.
                                                    colorID = colIdsWithThatColor.First();

                                                }
                                                return colorID;
                                            });
                                        var ak_des = Tabbles_decl.atom_key.NewAsk(Tabbles_decl.atom_set_key.NewAskLabel(catName));
                                        //myLog.printDisp("debug_outlook_plugin: creating tabble " + catName);
                                        db.dbCreateTabble2(ak_des, colorId, FSharpOption<string>.None, true, FSharpList<string>.Empty);
                                    }
                                }

                            }

                            // 2. tag the ct with those tabbles. The ct does not need to exist in tabbles'db.
                            {
                                var ct2 = Tabbles_decl.ct_key2.NewCt2Gen(new Tabbles_decl.generic_key2(x.gen.commandLine, x.gen.name));
                                var ct2Arr = new Tabbles_decl.ct_key2[] { ct2 };

                                //myLog.printDisp("debug_outlook_plugin: tagging email " + x.gen.name  );
                                db.tagCtsNonRecursive(
                                    SeqModule.ToList(ct2Arr),
                                    SeqModule.ToList(newTags),
                                    db.wakeUpThreads.DoNotWakeUpThreads);
                            }
                            // 3. in tabbles'db, remove all other tabbles from that ct

                            {
                                var cd = db.getDataOfCt(db.mergedLocal.LocalDb, ct); // this will not fail because the ct surely exists in tabbles'db right now
                                var cdg = (Tabbles_decl.ct_data.Cd_gen)cd;
                                var currentTags = cdg.Item.gd_tabbles;
                                var tagsToRemove = (from t in currentTags
                                                    where !(newTags.Contains(t))
                                                    select t).ToList();

                                var ctArr = new Tabbles_decl.ct_key[] { ct };
                                //myLog.printDisp("debug_outlook_plugin: untagging email " + x.gen.name);
                                db.untagCtsNonRecursive2(SeqModule.ToList(tagsToRemove), SeqModule.ToList(ctArr), false); // qui crasha perché non esiste il ct nel db TODO maurizio
                            }


                            // old code:
                            //if (db.ctExists(db.mergedLocal.LocalDb, ct)) // otherwise we should add the ct to the db. For now we are doing nothing if the email is not in tabbles'db
                            //{
                            //    var cd = db.getDataOfCt(db.mergedLocal.LocalDb, ct);
                            //    var cdg = (Tabbles_decl.ct_data.Cd_gen)cd;
                            //    var oldTags = cdg.Item.gd_tabbles;

                            //    // 1. in tabbles' db, remove the tabbles that should not be there for that email
                            //    {
                            //        var tagsToRemove = (from t in oldTags
                            //                            where !(newTags.Contains(t))
                            //                            select t).ToList();

                            //        var ctArr = new Tabbles_decl.ct_key[] { ct };
                            //        db.untagCtsNonRecursive2(SeqModule.ToList(tagsToRemove), SeqModule.ToList(ctArr), false); // qui crasha perché non esiste il ct nel db TODO maurizio
                            //    }


                            //    // 2. tag the emails with the tags they currently have in outlook (which are assumed to exist in Tabbles too!)
                            //    {
                            //        var ct2 = Tabbles_decl.ct_key2.NewCt2Gen(new Tabbles_decl.generic_key2(x.gen.commandLine, x.gen.name));
                            //        var ct2Arr = new Tabbles_decl.ct_key2[] { ct2 };
                            //        db.tagCtsNonRecursive(
                            //            SeqModule.ToList(ct2Arr),
                            //            SeqModule.ToList(newTags),
                            //            db.wakeUpThreads.DoNotWakeUpThreads);
                            //    }
                            //}
                        });
                    }
                    else if (messageObj is INeedToOpenSearch)
                    {
                        TabblesApi.API.ExecuteActionInGuiThread(() =>
                        {

                            Tabbles_logic.showQuickOpenTabblesDialogAsync(
                                FSharpOption<string>.None,
                                FSharpOption<System.Windows.Window>.None,
                                FSharpOption<WpfControlLibrary1.wnd_file_panel>.None,
                                ((result, aks) =>
                                {
                                    // let us specify what must happen after the user closes the quick-open-tabbles dialog and has chosen some tags

                                    if (result.IsQdr_canceled || result.IsQdr_menu)
                                    {
                                        // do nothing, the user canceled 
                                    }
                                    else
                                    {
                                        var aksWhichAreNotNormalTags =
                                            (from ak in aks
                                             where ak.IsAsk
                                             let ask = ((Tabbles_decl.atom_key.Ask)ak).Item
                                             where !(ask.IsAskLabel)
                                             select ask).ToArray();
                                        if (aksWhichAreNotNormalTags.Length > 0)
                                        {
                                            TabblesApi.API.ExecuteActionInGuiThread(() =>
                                            {
                                                popup.showPopupNoLists("Cannot execute", "With the Tabbles addin for Outlook, you can only open ordinary tabbles, not folder-shortcuts or other special tabbles.", Tabbles_decl.popupKind.PkNormal, popup.color.CError);
                                            });
                                        }
                                        else
                                        {
                                            var tags = (from ak in aks
                                                        where ak.IsAsk
                                                        let ask = ((Tabbles_decl.atom_key.Ask)ak).Item
                                                        where ask.IsAskLabel
                                                        let lab = (Tabbles_decl.atom_set_key.AskLabel)ask
                                                        select lab.Item).ToArray();
                                            SendMessageToOutlookBlocking(new MsgOpenMailsWithTags()
                                            {
                                                tags = tags
                                            });
                                        }
                                    }
                                }),
                                Misc.QuickOpenDisableDropDownMenu.Qod_disableDropdownMenu);

                        });
                    }
                    else if (messageObj is INeedToPingTabbles)
                    {
                        //a ping from add-in

                        // todo maurizio: probably we just need to check if the pipe is working
                    }
                    //else if (messageObj is INeedToRemoveTabblesFromCategories)
                    //{
                    //    myLog.printDisp("outlook system: tabbles plugin: message was: INeedToRemoveTabblesFromCategories...");
                    //    var x = (INeedToRemoveTabblesFromCategories)messageObj;
                    //}
                    //else if (messageObj is INeedToAddTabblesToCategories)
                    //{
                    //    myLog.printDisp("outlook system: tabbles plugin: message was: INeedToAddTabblesToCategories...");
                    //    var x = (INeedToAddTabblesToCategories)messageObj;
                    //}
                    else
                    {
                        TabblesApi.API.ExecuteActionInGuiThread(() =>
                        {
                            popup.showPopupNoLists("Error in plugin", "Outlook plugin sent a message not understood by the Tabbles plugin.", Tabbles_decl.popupKind.PkNormal, popup.color.CError);
                        });
                    }
                }
            }
        }

        public void onCtsTagged(IEnumerable<Tabbles_decl.ct_key2> cts, IEnumerable<string> tags)
        {
            var gens = (from ct in cts
                        where ct.IsCt2Gen
                        let gen2 = ((Tabbles_decl.ct_key2.Ct2Gen)ct).Item
                        select gen2.gk2_key).ToList();

            var tagsWithColors = new List<NameColorPair>(); // we are now going to fill this list
            // fill the list tagsWithColors 
            TabblesApi.API.LockDbA((() =>
                {
                    foreach (var t in tags)
                    {
                        // convert the tabble name to a "formula". (First convert the name to an atom, then the atom to a formula).
                        var fk = Tabbles_decl.formula_key.FkAtomic.NewFkAtomic(Tabbles_decl.atom_key.NewAsk(Tabbles_decl.atom_set_key.NewAskLabel(t)));
                        //get the data of the formula
                        var fd = db.getDataOfFk(db.mergedLocal.LocalDb, fk);
                        var ad = (Misc.formula_data.FdAtomic)fd;
                        var adl = (Tabbles_decl.atom_data.AdLabel)ad.Item;
                        var colorId = adl.Item.adl_cat;

                        // extract the color given the color id
                        var cd = db.getDataOfColor(db.mergedLocal.LocalDb, colorId);
                        // add the pair name/color to the list
                        tagsWithColors.Add(new NameColorPair
                        {
                            Color = cd.cad_cat_color,
                            Name = t
                        });
                    }
                }
            ));


            if (gens.Any())
            {
                SendMessageToOutlookBlocking(new MsgGensTagged
                {
                    gens = gens,
                    tags = tagsWithColors
                });
            }
        }

        public void onCtsUntagged(IEnumerable<Tabbles_decl.ct_key> cts, IEnumerable<string> tags)
        {
            var gens = (from ct in cts
                        where ct.IsCtGen
                        let gen2 = ((Tabbles_decl.ct_key.CtGen)ct).Item
                        select gen2).ToList();

            if (gens.Any())
            {
                SendMessageToOutlookBlocking(new MsgGensUntagged
                {
                    gens = gens,
                    tags = tags
                });
            }
        }

        //public void onTagsSelected(IEnumerable<string> tags)
        //{
        //    SendMessageToOutlookBlocking(new MsgOpenMailsWithTags()
        //    {
        //        tags = tags
        //    });
        //}

        public void onAksDeleted(IEnumerable<Tabbles_decl.atom_key> aks)
        {
            var aksStr = (from ak in aks
                          select db.stringOfAkForTabble(db.whatNameForFolderTabble.WhatNameForFolderTabble_fileName, ak)).ToList();
            SendMessageToOutlookBlocking(new MsgAtomKeysDeleted
            {
                AtomKeys = aksStr
            });
        }

        public void onAkCreated(Tabbles_decl.atom_key ak)
        {
            var name = db.stringOfAkForTabble(db.whatNameForFolderTabble.WhatNameForFolderTabble_fileName, ak);

            // we have to read the color of the given atom_key (abbreviated ak). Colors are stored by tabbles in a dictionary, where the key 
            // is a formula_key (abbreviated fk), and the value is a formula_data, which contains the color. 
            // Since we have an ak, not a fk, we first need to convert the ak to a fk. 

            var color = TabblesApi.API.LockDbF<string>((() =>
                 {
                     //convert the ak to an fk
                     var fk = Tabbles_decl.formula_key.FkAtomic.NewFkAtomic(ak);

                     //get the data of the fk
                     var fd = db.getDataOfFk(db.mergedLocal.LocalDb, fk);
                     var ad = (Misc.formula_data.FdAtomic)fd;
                     var adl = (Tabbles_decl.atom_data.AdLabel)ad.Item;
                     var colorId = adl.Item.adl_cat;

                     // extract the color given the color id
                     var cd = db.getDataOfColor(db.mergedLocal.LocalDb, colorId);
                     return cd.cad_cat_color;
                 }));

            SendMessageToOutlookBlocking(new MsgAtomKeyCreated
            {
                AtomKeyName = name,
                AtomKeyColor = color
            });
        }

        public string Name
        {
            get
            {
                return "Tabbles plugin for Outlook";
            }
        }

        public string PluginVersion
        {
            get
            {
                return "0.1";
            }
        }

        public string RequiredTabblesVersion
        {
            get
            {
                return "unspecified";
            }
        }

        public string Author
        {
            get
            {
                return "Maurizio Colucci - Yellow blue soft";
            }
        }


    }
}
