using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EmailMerge
{
    public class PSTFile
    {
        /// <summary>
        /// http://stackoverflow.com/questions/577904/can-i-read-an-outlook-2003-2007-pst-file-in-c
        /// </summary>


        private Application _outlookApp;
        private NameSpace _outlookNameSpace;
        public struct MailItemObject
        {
            public string FolderName;
            public MailItem MailObject;
        }

        public ConcurrentDictionary<string, MailItemObject> AllMailItems = new ConcurrentDictionary<string, MailItemObject>();
        public ConcurrentDictionary<string, MailItemObject> DuplicateMailItems = new ConcurrentDictionary<string, MailItemObject>();
        public ConcurrentDictionary<string, MailItemObject> OriginalMailItems = new ConcurrentDictionary<string, MailItemObject>();

        public List<MailItemObject> GetAllMailItems()
        {
            return AllMailItems.Select(keyvalue => keyvalue.Value).ToList();
        }

        public enum MailCompareType
        {
            In1NotIn2,
            In2NotIn1,
            InBoth
        }

        private FileInfo PSTFileInfo()
        {
            return new FileInfo(PSTFilePath);
        }

        public string GetPSTFilename()
        {
            return PSTFileInfo().Name;
        }

        public string GetPSTFileDirectory()
        {
            return PSTFileInfo().DirectoryName + @"\";
        }

        public readonly string PSTFilePath;
        public readonly string PSTName;

        public static string[] DefaultFoldersToIgnore = new[]{"Outbox",
                                                                "Calendar",
                                                                "Deleted Items",
                                                                "Junk E-mail",
                                                                "Contacts",
                                                                "Journal",
                                                                "Notes",
                                                                "Tasks",
                                                                "Drafts",
                                                                "RSS Feeds",
                                                                "Suggested Contacts",
                                                                "Conversation Action Settings",
                                                                "Quick Step Settings",
                                                                "News Feed"};


        #region CtorDtor

        public PSTFile(string pstFilePath, string pstName = null)
        {
            Console.WriteLine("Loading PST file: '" + pstFilePath + "'");
            _outlookApp = new Application();
            _outlookNameSpace = _outlookApp.GetNamespace("MAPI");

            PSTFilePath = pstFilePath;
            PSTName = pstName;

            LoadPSTFile();
            SetPSTName();
        }

        ~PSTFile()
        {
            RemovePSTFile();
        }

        public void RemovePSTFile()
        {
            Console.WriteLine("Removing PST file: '" + PSTName + "'");
            // Remove PST file from Default Profile
            if(CheckStoreExists(PSTName))
                _outlookNameSpace.RemoveStore(_outlookNameSpace.Stores[PSTName].GetRootFolder());
            _outlookApp.Quit();
            _outlookNameSpace = null;
            _outlookApp = null;
        }

        private bool CheckStoreExists(string storeName)
        {
            try
            {
                Store check = _outlookNameSpace.Stores[storeName];
                string checkName = check.DisplayName;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void LoadPSTFile()
        {
            // Add PST file (Outlook Data File) to Default Profile
            _outlookNameSpace.AddStoreEx(PSTFilePath, OlStoreType.olStoreDefault);
        }

        private void SetPSTName()
        {
            MAPIFolder objFolder = _outlookNameSpace.Folders.GetLast();
            objFolder.Name = PSTName ?? PSTFilePath;
        }

        #endregion


        #region Static Methods

        public List<MailItemObject> ComparePSTFiles(PSTFile pstFile2,
                                                MailCompareType compareType = MailCompareType.In1NotIn2)
        {
            return ComparePSTFiles(this, pstFile2, compareType);
        }

        public static List<MailItemObject> ComparePSTFiles(PSTFile pstFile1,
                                                            PSTFile pstFile2,
                                                            MailCompareType compareType = MailCompareType.In1NotIn2)
        {
            List<MailItemObject> compareRes = new List<MailItemObject>();

            switch (compareType)
            {
                case MailCompareType.In1NotIn2:
                    foreach (KeyValuePair<string, MailItemObject> mailItem in pstFile1.AllMailItems)
                    {

                        if (!pstFile2.AllMailItems.ContainsKey(mailItem.Key))
                            compareRes.Add(mailItem.Value);
                    }
                    break;
                case MailCompareType.In2NotIn1:
                    foreach (KeyValuePair<string, MailItemObject> mailItem in pstFile2.AllMailItems)
                    {

                        if (!pstFile1.AllMailItems.ContainsKey(mailItem.Key))
                            compareRes.Add(mailItem.Value);
                    }
                    break;

                case MailCompareType.InBoth:
                    foreach (KeyValuePair<string, MailItemObject> mailItem in pstFile1.AllMailItems)
                    {

                        if (pstFile2.AllMailItems.ContainsKey(mailItem.Key))
                            compareRes.Add(mailItem.Value);
                    }
                    break;
            }
            return compareRes;
        }

        public static void CreatePSTFile(string fileName, OlStoreType version)
        {
            if (!File.Exists(fileName))
            {
                Application _App = new Application();
                NameSpace _NameSpace = _App.GetNamespace("MAPI");
                _NameSpace.AddStoreEx(fileName, version);
                _App.Quit();
                _NameSpace = null;
                _App = null;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="pstFile1"></param>
        /// <param name="pstFile2"></param>
        /// <param name="mergedFilename"></param>
        /// <param name="duplicatesFilename1"></param>
        /// <param name="duplicatesFilename2"></param>
        /// <param name="saveDuplicates"></param>
        /// <param name="foldersToIgnore"></param>
        public static void MergePSTFiles(string directoryToMerge, string mergedFilename = null,
                                            string duplicatesFilename1 = null, string duplicatesFilename2 = null, bool saveDuplicates = true,
                                            string[] foldersToIgnore = null)
        {
            
            
            PSTFile mergedPSTFile = new PSTFile(mergedFilename, "Merged");

            var ext = new List<string> { "pst" };
            string[] filePaths = Directory.GetFiles(directoryToMerge, "*.*", SearchOption.AllDirectories).Where(s => ext.Contains(Path.GetExtension(s))).ToArray();

            foreach (string s in filePaths)
            {
                string fileName = Path.GetFileName(s);
                PSTFile sourceFile = new PSTFile(s, fileName);
                sourceFile.GetMailItems(true, foldersToIgnore);
                List<MailItemObject> mailItems = sourceFile.GetAllMailItems();
                Console.WriteLine("Mail items (" + mailItems.Count + ") loaded from:" + sourceFile.PSTFilePath);

                mergedPSTFile.AddMailItems(mailItems);
            }


            //Console.WriteLine("Merging PST:" + pstFile2.PSTFilePath);
            //pstFile2.GetMailItems(true, foldersToIgnore);


            //List<MailItemObject> mailItems2 = pstFile2.GetAllMailItems();
            //Console.WriteLine("Mail items (" + mailItems2.Count + ") loaded from:" + pstFile2.PSTFilePath);


            //mergedPSTFile.AddMailItems(mailItems2);

        }


        #endregion


        #region Add Mail Items

        //https://social.msdn.microsoft.com/Forums/vstudio/en-US/3dd2bd06-5738-4fb2-b628-0d7ab2be8157/how-to-directly-copy-a-mailitem-into-public-folder-sth-like-mailitemcopydestinationfolder?forum=vsto
        public void AddMailItems(List<MailItemObject> mailItems)
        {
            int count = 0;

            Parallel.ForEach(mailItems, mailItemObject =>
            {
                AddMailItems(mailItemObject.MailObject, mailItemObject.FolderName);
                count++;
            });
            Console.WriteLine("Added " + count + " mail items.");
        }
        public void AddMailItems(List<MailItem> mailItems, string folderName = "")
        {
            Parallel.ForEach(mailItems, mailItem =>
            {
                AddMailItems(mailItem, folderName);
            });
        }
        public void AddMailItems(MailItem mailItem, string folderName = "")
        {
            folderName = folderName ?? "Temp_" + PSTName;

            MAPIFolder destFolder = GetFolder(folderName);

            MailItem tempMailItem = mailItem.Copy();

            Console.WriteLine("Copying:" + tempMailItem.SentOn.ToString("dd-MM-yyyy_HH:mm:ss") + " :::: " + tempMailItem.Subject);

            tempMailItem.Move(destFolder);
        }

        public void AddFolder(string folderName)
        {
            lock (_outlookNameSpace)
            {
                if (!FolderExists(folderName))
                {
                    _outlookNameSpace.Stores[PSTName].GetRootFolder().Folders.Add(folderName);
                }
            }
        }

        private MAPIFolder GetFolder(string FolderPath, bool addIfMissing = true)
        {
            lock (_outlookNameSpace)
            {
                string[] folders = FolderPath.TrimStart(new[] { '\\' }).Split('\\');
                MAPIFolder currentFolder = _outlookNameSpace.Stores[PSTName].GetRootFolder();

                foreach (string folder in folders)
                {
                    if (!FolderExists(currentFolder, folder) && addIfMissing)
                        currentFolder.Folders.Add(folder);

                    currentFolder = currentFolder.Folders[folder];
                }
                return currentFolder;
            }

        }

        #endregion


        #region Save Duplicates

        public void SaveDuplicatesToFile(string pstFilename,
                                        string pstName = "DuplicateEmails",
                                        string duplicateFolder = "Duplicates",
                                        string originalFolder = "Originals")
        {
            Console.WriteLine("Saving duplicates to file: '" + pstFilename + "'");
            PSTFile pstDuplicates = new PSTFile(pstFilename, pstName);

            if (duplicateFolder != "" && DuplicateMailItems.Any())
                pstDuplicates.AddMailItems(DuplicateMailItems.Select(mailItem => mailItem.Value.MailObject).ToList(), duplicateFolder);
            if (originalFolder != "" && OriginalMailItems.Any())
                pstDuplicates.AddMailItems(OriginalMailItems.Select(mailItem => mailItem.Value.MailObject).ToList(), originalFolder);
        }

        #endregion


        #region Get Mail Items

        public void GetMailItems(bool includeSubFolders = true, string[] foldersToIgnore = null)
        {
            AllMailItems.Clear();
            DuplicateMailItems.Clear();
            OriginalMailItems.Clear();


            Microsoft.Office.Interop.Outlook.Stores stores = _outlookNameSpace.Stores;
            foreach (Microsoft.Office.Interop.Outlook.Store store in stores)
            {
                try { 
                if (store.IsDataFileStore == true)
                {


                    if(store.FilePath.Contains(this.PSTName))
                    {

                            Console.WriteLine(String.Format("Loading Store: "
                            + store.DisplayName
                            + "\n" + "File Path: "
                            + store.FilePath + "\n"));

                            MAPIFolder rootFolder = store.GetRootFolder();
                            GetMailItems(rootFolder.Folders, includeSubFolders, foldersToIgnore);
                            break;
                    }
                    

                        Console.WriteLine(String.Format("Skipping Store: "
                        + store.DisplayName
                        + "\n" + "File Path: "
                        + store.FilePath + "\n"));

                    }
                } catch(System.Exception e)
                {
                    throw new System.Exception("Problem loading Store", e);
                }
            }


        }

        private void GetMailItems(Folders folders, bool includeSubFolders = true, string[] foldersToIgnore = null)
        {
            foreach (Folder folder in folders)
            {
                GetMailItems(folder, includeSubFolders, foldersToIgnore);
            }
            Console.WriteLine("Finished loading Mail Items for folders, num items: " + AllMailItems.Count);
        }


        private void GetMailItems(Folder folder, bool includeSubFolders = true, string[] foldersToIgnore = null)
        {
            try
            {
                if (foldersToIgnore != null && foldersToIgnore.Contains(folder.Name)) return;

                if (includeSubFolders)
                {
                    //Calls recursively until all subfolders covered
                    foreach (Folder subFolder in folder.Folders)
                    {
                        GetMailItems(subFolder, includeSubFolders);
                    }
                }

                string folderPath = folder.FolderPath.Replace(@"\\" + PSTName, "");
                Console.WriteLine("Loading Mail Items for folder: '" + folder.Name + "' in path: '" + folderPath + "'"+" from "+PSTFilePath);
                int i = 0;

                // Traverse through all folders in the PST file

                Items items = folder.Items;

                Parallel.ForEach(items.Cast<object>(), item =>
                {

                    if (item is MailItem)
                    {
                        MailItem mailItem = item as MailItem;
                        string key = MailItemKey(mailItem);

                        if (AllMailItems.ContainsKey(key))
                        {
                            DuplicateMailItems.GetOrAdd(key + i++, new MailItemObject() { MailObject = mailItem, FolderName = folderPath });
                            if (!OriginalMailItems.ContainsKey(key))
                                OriginalMailItems.GetOrAdd(key, AllMailItems[key]);
                        }
                        else
                        {
                            AllMailItems.GetOrAdd(key, new MailItemObject() { MailObject = mailItem, FolderName = folderPath });
                        }
                    }
                }
                );

            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        #endregion


        #region MailItem Utils

        public string MailItemKey(MailItem mailItem)
        {
            try
            {

                string key = mailItem.SenderEmailAddress + "|"
                            //+ String.Concat(GetSMTPAddressForRecipients(mailItem)) + "|"
                            + mailItem.SentOn.ToString("dd-MM-yyyy_HH:mm:ss") + "|"
                            + mailItem.Subject + "|"
                            + GetMailBodyHashCode(mailItem);// + "|"
                            //+ GetAttachmentsHashCode(mailItem);

                return key;
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.InnerException.Message);
                throw e;
            }
        }

        private string GetMailBodyHashCode(MailItem mailItem)
        {
            if (mailItem.Body != null)
                return mailItem.Body.GetHashCode().ToString();
            return null;
        }

        private string GetAttachmentsHashCode(MailItem mailItem)
        {
            string attachmentKey = "";

            foreach (Attachment attachment in mailItem.Attachments)
            {
                attachmentKey += attachment.FileName.GetHashCode();
            }
            return attachmentKey;
        }

        private List<string> GetSMTPAddressForRecipients(MailItem mailItem)
        {
            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            Recipients recips = mailItem.Recipients;

            List<string> recipEmails = new List<string>();

            foreach (Recipient recip in recips)
            {
                PropertyAccessor pa = recip.PropertyAccessor;
                string smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS).ToString();

                recipEmails.Add(smtpAddress);
            }

            return recipEmails;
        }

        private bool FolderExists(string folderName)
        {
            return FolderExists(_outlookNameSpace.Stores[PSTName].GetRootFolder(), folderName);
        }

        private bool FolderExists(MAPIFolder baseFolder, string folderName)
        {
            try
            {
                MAPIFolder testFolder = baseFolder.Folders[folderName];
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

    }
}