using EPDM.Interop.epdm;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace BlueByte.SOLIDWORKS.PDMProfessional.Extensions
{

    /// <summary>
    /// Batch get settings for <see cref="Extensions.Extension.GetFiles(IEdmFolder5, BatchGetFilesSettings, string[])"/>
    /// </summary>
    public struct BatchGetFilesSettings
    {

        /// <summary>
        /// Gets or sets the file extensions. Must include dot before extension.
        /// </summary>
        /// <value>
        /// The file extensions.
        /// </value>
        public string[] FileExtensions { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="BatchGetFilesSettings"/> is enabled.
        /// </summary>
        /// <value>
        ///   <c>true</c> if enabled; otherwise, <c>false</c>.
        /// </value>
        public bool Enabled { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="BatchGetFilesSettings"/> is recursive.
        /// </summary>
        /// <value>
        ///   <c>true</c> if recursive; otherwise, <c>false</c>.
        /// </value>
        public bool Recursive { get; set; }
        /// <summary>
        /// Gets or sets the search object.
        /// </summary>
        /// <value>
        /// The search object.
        /// </value>
        public IEdmSearch5 SearchObject { get; set; }
    }

    /// <summary>
    /// Action required by the user depending on the state of <see cref="IEdmFile5"/>.
    /// </summary>
    public enum CheckoutAction
    {
        /// <summary>
        /// No action required. File is checked out by you.
        /// </summary>
        DoNothingFileCheckedOutByMe,

        /// <summary>
        /// File is checked in and can be checked out.
        /// </summary>
        FileCheckedInCanBeCheckedOut,

        /// <summary>
        /// File is checked out by someone else.
        /// </summary>
        CheckedOutBySomeoneElse
    }
    

    



    /// <summary>
    /// Comparison between the local copy and the server copy.
    /// </summary>
    public enum LocalCopyState
    {
        /// <summary>
        /// There is no local copy in your vault view.
        /// </summary>
        ThereIsNoLocalCopy,

        /// <summary>
        /// Local copy is latest.
        /// </summary>
        LocalAndServerCopySynced,

        /// <summary>
        /// New changes were pushed to server. Your local copy is obselete.
        /// </summary>
        LocalCopyObseleteServerHasNewOne,

        /// <summary>
        /// Your local copy is not the latest one.
        /// </summary>
        LocalCopyIsBehindServer,
    }

    /// <summary>
    /// SOLIDWORKS document type for <see cref="IEdmFile5"/>.
    /// </summary>
    public enum swPDMDocumentType_e
    {
        /// <summary>
        /// Any file that is not a SOLIDWORKS file.
        /// </summary>
        Generic,

        /// <summary>
        /// SOLIDWORTKS part document.
        /// </summary>
        Part,

        /// <summary>
        /// SOLIDWORKS assembly document.
        /// </summary>
        Assembly,

        /// <summary>
        /// SOLIDWORKS drawing document.
        /// </summary>
        Drawing,
    }


    /// <summary>
    /// PDM extensions
    /// </summary>
    public static class Extension
    {
        #region Public Methods


        /// <summary>
        /// Batch gets variable values.
        /// </summary>
        /// <param name="items">The items. May include files and folders</param>
        /// <param name="variableNames">The variables names.</param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public static Dictionary<IEdmObject5, Dictionary<string, object>> BatchGetVariableValues(this IEdmObject5[] items, string[] variableNames)
        {
            throw new NotImplementedException();
        }

        


        /// <summary>
        /// Gets the name of the previous file's state.
        /// </summary>
        /// <param name="file">File</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">File is null.</exception>
        /// <exception cref="System.Runtime.InteropServices.COMException">
        /// File has state changes history.
        /// or
        /// File has no history.
        /// </exception>
        public static string GetPreviousState(this IEdmFile5 file)
        {

            var stateName = string.Empty;

            if (file == null)
                throw new ArgumentNullException("file");

            var vault = file.Vault as IEdmVault7;

            var history = vault.CreateUtility(EdmUtility.EdmUtil_History) as IEdmHistory;

            EdmHistoryItem[] historyitems = null;
            history.AddFile(file.ID);
            history.GetHistory(ref historyitems, (int)EdmHistoryType.Edmhist_FileState);

            if (history == null)
                throw new COMException("File has state changes history");

            if (historyitems.Length == 0)
                throw new COMException("File has no history.");

          

            stateName = historyitems[0].moData.mbsStrData1;

            var workflow = vault.GetObject(EdmObjectType.EdmObject_Workflow, (file.CurrentState as IEdmState6).WorkflowID) as IEdmWorkflow6;

            

            return stateName.Replace(workflow.Name,"").Trim();

        }

        /// <summary>
        /// Batch check in an array of files.
        /// </summary>
        /// <param name="Vault">The vault.</param>
        /// <param name="files">The files.</param>
        /// <param name="locks">The locks.</param>
        /// <param name="Handle">The handle.</param>
        /// <param name="EdmUnlockOpCallback">The edm unlock op callback.</param>
        /// <returns></returns>
        public static bool BatchCheckIn(this IEdmVault5 Vault, IEdmFile5[] files, EdmUnlockBuildTreeFlags locks, int Handle, IEdmUnlockOpCallback EdmUnlockOpCallback = null)
        {
            int intlocks;

            if (locks == 0)
                intlocks = (int)EdmUnlockBuildTreeFlags.Eubtf_ShowCloseAfterCheckinOption + (int)EdmUnlockBuildTreeFlags.Eubtf_MayUnlock;
            else
                intlocks = (int)locks;

            // create batch unlocker object
            var batchUnlocker = (IEdmBatchUnlock2)(Vault as IEdmVault8).CreateUtility(EdmUtility.EdmUtil_BatchUnlock) as IEdmBatchUnlock2;
            // create the selection list
            var list = new List<EdmSelItem>();
            var selectedFile = default(EdmSelItem);
            foreach (var file in files)
            {
                var ppoRetParentFolder = default(IEdmFolder5);
                string localPath = file.GetLocalPath(file.GetParentFolderID());
                var aFile = Vault.GetFileFromPath(localPath, out ppoRetParentFolder);
                selectedFile = new EdmSelItem();
                IEdmPos5 aPos = aFile.GetFirstFolderPosition();
                IEdmFolder5 aFolder = aFile.GetNextFolder(aPos);
                selectedFile.mlDocID = aFile.ID;
                selectedFile.mlProjID = aFolder.ID;
                list.Add(selectedFile);
            }

            var vault1 = new EdmVault5();
            vault1 = (EdmVault5)Vault;
            // do not change this - the array must init as null
            EdmSelItem[] ppoSelection = null;

            ppoSelection = list.ToArray();
            batchUnlocker.AddSelection(vault1, ref ppoSelection);

            //create tree
            batchUnlocker.CreateTree(Handle, intlocks);
            // unlock file
            batchUnlocker.UnlockFiles(Handle, EdmUnlockOpCallback);
            return true;
        }

        /// <summary>
        /// Batch check-out an array of PDM files.
        /// </summary>
        /// <param name="Vault">Vault object.</param>
        /// <param name="files">array of filenames.</param>
        /// <param name="Handle">Handle of the parent window.</param>
        /// <returns>True if successful, false if not.</returns>
        /// <remarks>Method is not thread-safe. Use this method only in the main thread of your application.</remarks>
        public static bool BatchCheckIn(this IEdmVault5 Vault, string[] files, int Handle, EdmUnlockBuildTreeFlags locks, IEdmUnlockOpCallback EdmUnlockOpCallback = null)
        {
            // create batch unlocker object
            var batchUnlocker = (IEdmBatchUnlock2)(Vault as IEdmVault8).CreateUtility(EdmUtility.EdmUtil_BatchUnlock) as IEdmBatchUnlock2;
            // create the selection list
            var list = new List<EdmSelItem>();
            var selectedFile = default(EdmSelItem);
            foreach (var file in files)
            {
                var ppoRetParentFolder = default(IEdmFolder5);
                var aFile = Vault.GetFileFromPath(file, out ppoRetParentFolder);
                selectedFile = new EdmSelItem();
                IEdmPos5 aPos = aFile.GetFirstFolderPosition();
                IEdmFolder5 aFolder = aFile.GetNextFolder(aPos);
                selectedFile.mlDocID = aFile.ID;
                selectedFile.mlProjID = aFolder.ID;
                list.Add(selectedFile);
            }

            var vault1 = new EdmVault5();
            vault1 = (EdmVault5)Vault;
            // do not change this - the array must init as null
            EdmSelItem[] ppoSelection = null;

            ppoSelection = list.ToArray();
            batchUnlocker.AddSelection(vault1, ref ppoSelection);

            int intlocks;

            if (locks == 0)
                intlocks = (int)EdmUnlockBuildTreeFlags.Eubtf_ShowCloseAfterCheckinOption + (int)EdmUnlockBuildTreeFlags.Eubtf_MayUnlock;
            else
                intlocks = (int)locks;

            //create tree
            batchUnlocker.CreateTree(Handle, intlocks);
            // unlock files
            batchUnlocker.UnlockFiles(Handle, EdmUnlockOpCallback);
            return true;
        }

        /// <summary>
        /// Batch check out (or get versions of) an array of PDM files.
        /// </summary>
        /// <param name="Vault">Vault object.</param>
        /// <param name="files">array of IEdmFile5.</param>
        /// <param name="getCmds">Get commands.</param>
        /// <param name="Handle">Handle of the parent window.</param>
        /// <returns>True if successful, false if not.</returns>
        /// <remarks>Method is not thread-safe. Use this method only in the main thread of your application.</remarks>
        public static bool BatchGetFiles(this IEdmVault5 Vault, IEdmFile5[] files, EdmGetCmdFlags getCmds, int Handle, IEdmGetOpCallback EdmGetOpCallback = null)
        {
            // create batch locker object
            var batchLocker = (IEdmBatchGet)(Vault as IEdmVault8).CreateUtility(EdmUtility.EdmUtil_BatchGet) as IEdmBatchGet;
            // create the selection list
            var list = new List<EdmSelItem>();
            var selectedFile = default(EdmSelItem);
            foreach (var file in files)
            {
                var ppoRetParentFolder = default(IEdmFolder5);
                string localPath = file.GetLocalPath(file.GetParentFolderID());
                var aFile = Vault.GetFileFromPath(localPath, out ppoRetParentFolder);
                selectedFile = new EdmSelItem();
                IEdmPos5 aPos = aFile.GetFirstFolderPosition();
                IEdmFolder5 aFolder = aFile.GetNextFolder(aPos);
                selectedFile.mlDocID = aFile.ID;
                selectedFile.mlProjID = aFolder.ID;
                list.Add(selectedFile);
            }

            var vault1 = new EdmVault5();
            vault1 = (EdmVault5)Vault;
            // do not change this - the array must init as null
            EdmSelItem[] ppoSelection = null;

            ppoSelection = list.ToArray();
            batchLocker.AddSelection(vault1, ref ppoSelection);

            //create tree
            batchLocker.CreateTree(Handle, (int)getCmds);
            // lock file
            batchLocker.GetFiles(Handle, EdmGetOpCallback);
            return true;
        }

        /// <summary>
        /// Batch check out (or get versions of) an array of PDM files.
        /// </summary>
        /// <param name="Vault">Vault object.</param>
        /// <param name="files">array of filenames.</param>
        /// <param name="Handle">Handle of the parent window.</param>
        /// <returns>True if successful, false if not.</returns>
        /// <remarks>Method is not thread-safe. Use this method only in the main thread of your application.</remarks>
        public static bool BatchGetFiles(this IEdmVault5 Vault, string[] files, int Handle, IEdmGetOpCallback EdmGetOpCallback = null)
        {
            // create batch locker object
            var batchLocker = (IEdmBatchGet)(Vault as IEdmVault8).CreateUtility(EdmUtility.EdmUtil_BatchGet) as IEdmBatchGet;
            // create the selection list
            var list = new List<EdmSelItem>();
            var selectedFile = default(EdmSelItem);
            foreach (var file in files)
            {
                var ppoRetParentFolder = default(IEdmFolder5);
                var aFile = Vault.GetFileFromPath(file, out ppoRetParentFolder);
                selectedFile = new EdmSelItem();
                IEdmPos5 aPos = aFile.GetFirstFolderPosition();
                IEdmFolder5 aFolder = aFile.GetNextFolder(aPos);
                selectedFile.mlDocID = aFile.ID;
                selectedFile.mlProjID = aFolder.ID;
                list.Add(selectedFile);
            }

            var vault1 = new EdmVault5();
            vault1 = (EdmVault5)Vault;
            // do not change this - the array must init as null
            EdmSelItem[] ppoSelection = null;

            ppoSelection = list.ToArray();
            batchLocker.AddSelection(vault1, ref ppoSelection);

            //create tree
            batchLocker.CreateTree(Handle, (int)EdmGetCmdFlags.Egcf_Lock);
            // lock file
            batchLocker.GetFiles(Handle, EdmGetOpCallback);
            return true;
        }

        /// <summary>
        /// Batch get local copies of an array of files.
        /// </summary>
        /// <param name="Vault"></param>
        /// <param name="files"></param>
        /// <param name="versions"></param>
        /// <param name="Handle"></param>
        /// <param name="EdmGetOpCallback"></param>
        /// <returns></returns>
        public static bool BatchGetLocalCopies(this IEdmVault5 Vault, IEdmFile5[] files, int[] versions, int Handle, IEdmGetOpCallback EdmGetOpCallback = null)
        {
            // create batch locker object
            var batchLocker = (IEdmBatchGet)(Vault as IEdmVault8).CreateUtility(EdmUtility.EdmUtil_BatchGet) as IEdmBatchGet;
            // create the selection list

            foreach (var file in files)
            {
                var version = versions[Array.IndexOf(files, file)];
                var v = Vault as EdmVault5;
                batchLocker.AddSelectionEx(v, file.ID, file.GetParentFolderID(), version);
            }
            //create tree
            batchLocker.CreateTree(Handle, (int)EdmGetCmdFlags.Egcf_Nothing);
            // lock file
            batchLocker.GetFiles(Handle, EdmGetOpCallback);
 
            return true;
        }

        /// <summary>
        /// Sets the variable value.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <param name="variableName">Name of the variable.</param>
        /// <param name="value">The value.</param>
        /// <param name="bOnlyIfPartOfCard">True to store the variable only if it is part of the file or folder data card, false to store the variable as hidden data if it is not part of the file or folder data card</param>
        /// <param name="configurationName">Name of the configuration.</param>
        /// <exception cref="ArgumentNullException">
        /// file
        /// or
        /// Configuration cannot be empty
        /// </exception>
        /// <exception cref="Exception">Failed to set variable.</exception>
        public static void SetVariableValue(this IEdmFile5 file, string variableName, object value, bool bOnlyIfPartOfCard = false, string configurationName = "@")
        {
            if (file == null)
                throw new ArgumentNullException(nameof(file));

            // generic file
            if (file.IsAssembly() ==false && file.IsDrawing() == false && file.IsPart() == false)
                configurationName = string.Empty;
            else
            {
                if (string.IsNullOrWhiteSpace(configurationName))
                    throw new ArgumentNullException($"{nameof(configurationName)} cannot be empty for SOLIDWORKS files.");
            }

                try
            {
                IEdmEnumeratorVariable8 fileEnumerator = file.GetEnumeratorVariable() as IEdmEnumeratorVariable8;
                fileEnumerator.SetVar(variableName, configurationName, value, bOnlyIfPartOfCard);
                fileEnumerator.CloseFile(true);
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to set variable.", ex);
            }
        }


        /// <summary>
        /// Sets the variable value.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="variableName">Name of the variable.</param>
        /// <param name="value">The value.</param>
        /// <param name="bOnlyIfPartOfCard">True to store the variable only if it is part of the file or folder data card, false to store the variable as hidden data if it is not part of the file or folder data card</param>
        /// <exception cref="ArgumentNullException">folder</exception>
        /// <exception cref="Exception">Failed to set variable.</exception>
        public static void SetVariableValue(this IEdmFolder5 folder, string variableName, object value, bool bOnlyIfPartOfCard = false)
        {
            if (folder == null)
                throw new ArgumentNullException(nameof(folder));

            
            try
            {
                IEdmEnumeratorVariable8 fileEnumerator = folder as IEdmEnumeratorVariable8;
                fileEnumerator.SetVar(variableName, string.Empty, value, bOnlyIfPartOfCard);
                fileEnumerator.CloseFile(true);
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to set variable.", ex);
            }
        }

        /// <summary>
        /// Gets the file's variable value
        /// </summary>
        /// <param name="file">The file.</param>
        /// <param name="variableName">Name of a variable.</param>
        /// <param name="configuration">File's configuration</param>
        /// <returns>File's variable value as an object. Throws an exception if getting a variable failed.</returns>
        public static object GetVariableValue(this IEdmFile5 file, string variableName, string configuration = "@")
        {
            IEdmEnumeratorVariable8 fileEnumerator = file.GetEnumeratorVariable() as IEdmEnumeratorVariable8;
            object variableValue = null;
            
                fileEnumerator.GetVar(variableName, configuration, out variableValue);
                return variableValue;
            
        }

        /// <summary>
        /// Gets all variables.
        /// </summary>
        /// <param name="vault">The vault.</param>
        /// <returns></returns>
        public static IEdmVariable5[] GetAllVariables(this IEdmVault5 vault)
        {
            var variableMgr = vault as IEdmVariableMgr5;

            var l = new List<IEdmVariable5>();

            var pos = variableMgr.GetFirstVariablePosition();


           while (pos.IsNull == false)
            {
                l.Add(variableMgr.GetNextVariable(pos));
            }


            return l.ToArray();

        }

        /// <summary>
        /// Gets all variable values.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <param name="variableNames">The variable ids.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">
        /// file
        /// or
        /// variableIds
        /// </exception>
        public static Dictionary<string,Dictionary<string, object>> BatchGetVariables(this IEdmFile5 file, string[] variableNames)
        {
            if (file == null)
                throw new ArgumentNullException(nameof(file));


            if (variableNames == null)
                throw new ArgumentNullException(nameof(variableNames));

 

   
           

            var configurationNames = file.GetConfigurationNames().ToList();
            configurationNames.Add("@");

            var ret = new Dictionary<string, Dictionary<string, object>>();


            var enm = file.GetEnumeratorVariable() as IEdmEnumeratorVariable8;

            foreach (var configurationName  in configurationNames)
            {
                ret.Add(configurationName, new Dictionary<string, object>());

                
                foreach (var variableName in variableNames)
                {
                    var value = new object();
                    enm.GetVarFromDb(variableName, configurationName, out value);
                    ret[configurationName].Add(variableName, value);
                }
 
               
      
            }


            enm.CloseFile(true);


            return ret;

        }



        /// <summary>
        /// Gets the file's variable value from the server.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <param name="variableName">Name of a variable.</param>
        /// <param name="configuration">File's configuration</param>
        /// <returns>File's variable value as an object. Throws an exception if getting a variable failed.</returns>
        public static object GetVariableFromDb(this IEdmFile5 file, string variableName, string configuration = "@")
        {
            IEdmEnumeratorVariable8 fileEnumerator = file.GetEnumeratorVariable() as IEdmEnumeratorVariable8;
            object variableValue = null;
           
            fileEnumerator.GetVarFromDb(variableName, configuration, out variableValue);
            return variableValue;
           
        }


        /// <summary>
        /// Gets the folder's variable value
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="variableName">Name of a variable.</param>
        /// <returns>Folder's variable value as an object. Throws an exception if getting a variable failed.</returns>
        public static object GetVariableValue(this IEdmFolder5 folder, string variableName)
        {
            var folderEnumerator = folder as IEdmEnumeratorVariable5;
            object variableValue = null;
            try
            {
                folderEnumerator.GetVar(variableName, string.Empty, out variableValue);
                return variableValue;
            }
            catch (Exception ex)
            {
                throw new Exception($"Could not get a variable: {ex.Message}");
            }
        }


        /// <summary>
        /// Get all the children items via the iterator pattern.
        /// </summary>
        /// <typeparam name="T">Children time type.</typeparam>
        /// <typeparam name="V">Parent item.</typeparam>
        /// <param name="manager"></param>
        /// <returns></returns>
        public static T[] GetAllChildren<T,V>(this V manager)    
        {
         

            var temporaryList = new List<T>();
            var name = Microsoft.VisualBasic.Information.TypeName(manager);
            switch (typeof(T).Name)
            {
                case nameof(IEdmUser5):
                    if (name.StartsWith("IEdmUserGroup"))
                    {
                        var group = (IEdmUserGroup5)manager;
                        var position = group.GetFirstUserPosition();
                        while (position.IsNull == false)
                        {
                            temporaryList.Add((T)group.GetNextUser(position));
                        }

                    }
                    break;
                case nameof(IEdmUserGroup5):
                    if (name == nameof(IEdmUserMgr5) || name.Contains("EdmVault"))
                    {
                        var userMgr5 = (IEdmUserMgr5)manager;
                        var position = userMgr5.GetFirstUserGroupPosition();
                        while (position.IsNull == false)
                        {
                            temporaryList.Add((T)userMgr5.GetNextUserGroup(position));
                        }
                    }
                    break;
                case nameof(IEdmFile5):
                    if (name.StartsWith("IEdmFolder"))
                    {
                        var folder = (IEdmFolder5)manager;
                        var position = folder.GetFirstFilePosition();
                        while (position.IsNull == false)
                        {
                            temporaryList.Add((T)folder.GetNextFile(position));
                        }
                    }
                    break;
                case nameof(IEdmFolder5):
                    if (name.StartsWith("IEdmFolder"))
                    {
                        var folder = (IEdmFolder5)manager;
                        var position = folder.GetFirstSubFolderPosition();
                        while (position.IsNull == false)
                        {
                            temporaryList.Add((T)folder.GetNextSubFolder(position));
                        }
                    }
                    break;
                case nameof(IEdmVariable5):

                    if (name.StartsWith("IEdmUserMgr") || name.Contains("EdmVault") || name == nameof(IEdmVault5) || name == nameof(EdmVault5Class))
                    {
                        var variableMgr = (IEdmVariableMgr5)manager;
                        var position = variableMgr.GetFirstVariablePosition();
                        while (position.IsNull == false)
                        {
                            temporaryList.Add((T)variableMgr.GetNextVariable(position));
                        }
                    }
                    break;
                case nameof(IEdmWorkflow5):
                    if (name == nameof(IEdmWorkflow5) || name == nameof(IEdmVault5) || name == nameof(EdmVault5Class))
                    {
                        var workflowMgr = (IEdmWorkflowMgr6)manager;
                        var position = workflowMgr.GetFirstWorkflowPosition();
                        while (position.IsNull == false)
                        {
                            temporaryList.Add((T)workflowMgr.GetNextWorkflow(position));
                        }
                    }
                    break;
                case nameof(IEdmTransition5):
                    if (name == nameof(IEdmTransition5))
                    {
                        var workflow = (IEdmWorkflow5)manager;
                        var position = workflow.GetFirstTransitionPosition();
                        while (position.IsNull == false)
                        {
                            temporaryList.Add((T)workflow.GetNextTransition(position));
                        }
                    }
                    break;
                case nameof(IEdmState5):
                    if (name == nameof(IEdmState5))
                    {
                        var workflow = (IEdmWorkflow5)manager;
                        var position = workflow.GetFirstStatePosition();
                        while (position.IsNull == false)
                        {
                            temporaryList.Add((T)workflow.GetNextState(position));
                        }
                    }
                    break;
                default:
                    break;
            }

            return temporaryList.ToArray();
        }


        /// <summary>
        /// Returns an array of all variable names.
        /// </summary>
        /// <param name="variableMgr">Variable Manager.</param>
        /// <returns></returns>
        public static string[] GetVariableNames(this IEdmVariableMgr5 variableMgr)
        {

            var position = variableMgr.GetFirstVariablePosition();
            var varList = new List<string>();

            while (!position.IsNull)
            {
                var nextVar = variableMgr.GetNextVariable(position);
                varList.Add(nextVar.Name);
            }

            return varList.ToArray();
        }
 
        /// <summary>
        /// Batch rollback files to specific versions.
        /// </summary>
        /// <param name="vault"></param>
        /// <param name="files">Files</param>
        /// <param name="versions">Versions</param>
        /// <param name="handle">Parent window handle</param>
        /// <returns></returns> 
        public static bool BatchRollbackToVersions(this IEdmVault5 vault, IEdmFile5[] files, int[] versions, int handle)
        {
            if (files.Length != versions.Length)
                throw new Exception("versions and files array do not match in length. Missing versions or files.");
            var callback = default(IEdmGetOpCallback);

            BatchGetFiles(vault, files, EdmGetCmdFlags.Egcf_Lock, handle, callback);

            BatchGetLocalCopies(vault, files, versions, handle);

            BatchCheckIn(vault, files, EdmUnlockBuildTreeFlags.Eubtf_MayUnlock, handle);

            return true;

        }


        /// <summary>
        /// Adds sub-folders with specified rights and permissions.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="relativePath">The relative path.</param>
        /// <param name="parentHandle">The parent handle.</param>
        /// <param name="rights">The rights.</param>
        /// <exception cref="System.ArgumentNullException">folder</exception>
        /// <exception cref="System.ArgumentException"></exception>
        public static void AddFolder2(this IEdmFolder5 folder, string relativePath, int parentHandle = 0, EdmFolderData rights = null)
        {
            if (folder == null)
                throw new ArgumentNullException(nameof(folder));

            if (string.IsNullOrWhiteSpace(relativePath))
                throw new ArgumentException($"{nameof(folder)} is empty or null.");


            var arr = relativePath.Split("\\".ToCharArray());

            if (arr != null)
            {
                var currentFolder = folder;

                foreach (var subfolder in arr)
                {
                    currentFolder = currentFolder.AddFolder(parentHandle, subfolder, rights);
                }
            }
        
        }


        /// <summary>
        /// Returns the value of a boolean variable.
        /// </summary>
        /// <param name="file">File.</param>
        /// <param name="variableName">Variable name.</param>
        /// <param name="configurationName">Configuration name.</param>
        /// <returns>true or false</returns>
        public static bool GetBooleanVariableValue(this IEdmFile5 file, string variableName, string configurationName = "")
        {
            if (string.IsNullOrWhiteSpace(variableName))
                throw new ArgumentNullException("variableName");

            variableName = variableName.Trim();

            var enumerator = file.GetEnumeratorVariable() as IEdmEnumeratorVariable10;
            var fileType = file.GetFileType();
            switch (fileType)
            {
                case swPDMDocumentType_e.Part:
                case swPDMDocumentType_e.Drawing:
                case swPDMDocumentType_e.Assembly:
                    if (string.IsNullOrWhiteSpace(configurationName))
                        configurationName = "@";
                    break;

                default:
                case swPDMDocumentType_e.Generic:
                    configurationName = string.Empty;
                    break;
            }

            object value = null;

            enumerator.GetVarFromDb(variableName, configurationName, out value);

            if (value != null)
            {
                if ((int)value == -1)
                {
                    enumerator.GetVar2(variableName, variableName, file.GetParentFolderID(), out value);
                }
            }
            else
            {
                throw new Exception($"Returned value is nothing. Please check if variable '{variableName}' exists.");
            }

            bool ret = false;
            if ((int)value == 1)

                ret = true;
            if ((int)value == 0)
                ret = false;
            enumerator.CloseFile(true);

            return ret;
        }

        /// <summary>
        /// Gets an array of configuration names.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="version">Version number</param>
        /// <returns></returns>
        public static string[] GetConfigurationNames(this IEdmFile5 file, int version = 0)
        {
            EdmStrLst5 configList;
            if (version != 0)
                configList = file.GetConfigurations();
            else
                configList = file.GetConfigurations(version);

            var pos = configList.GetHeadPosition();

            var l = new List<string>();
            while (pos.IsNull == false)
            {
                l.Add(configList.GetNext(pos));
            }

            l.Remove("@");

            return l.ToArray();
        }


        /// <summary>
        /// Gets files from a folder.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="BatchGetFiles">Batch settings. This is a structure.</param>
        /// <returns>Array of <see cref="IEdmFile5"/></returns>
        public static IEdmFile5[] GetFiles(this IEdmFolder5 folder, BatchGetFilesSettings BatchGetFiles = default(BatchGetFilesSettings))
        {
            if (folder == null)
                throw new ArgumentNullException(nameof(folder));

            if (BatchGetFiles.Enabled == false)
            {
                var l = new List<IEdmFile5>();
                var pos = folder.GetFirstFilePosition();

                while (pos.IsNull == false)
                {
                    var file = folder.GetNextFile(pos);
                    if (BatchGetFiles.FileExtensions != null)
                    {
                        if (BatchGetFiles.FileExtensions.FirstOrDefault(x => x.Trim().ToLower() == Path.GetExtension(file.Name).Trim().ToLower()) != null)
                        {
                            l.Add(file);
                        }
                    }
                    else
                        l.Add(file);
                }
                return l.ToArray();
            }
            else
            {
                var vault = folder.Vault as IEdmVault21;
                IEdmSearch9 searchFunction;
                if (BatchGetFiles.SearchObject != null)
                {
                    searchFunction = BatchGetFiles.SearchObject as IEdmSearch9;
                    searchFunction.Recursive = searchFunction.Recursive;
                }
                else
                {
                    searchFunction = vault.CreateSearch2() as IEdmSearch9;
                    searchFunction.FindFiles = true;
                    searchFunction.Recursive = BatchGetFiles.Recursive;


                }

                if (BatchGetFiles.FileExtensions == null || BatchGetFiles.FileExtensions.Length == 0)
                {
                    // do nothing
                }
                else
                {
                    BatchGetFiles.FileExtensions = BatchGetFiles.FileExtensions.ToList().Select(x => $"*{x}").ToArray();
                    var searchQ = string.Join(" | ", BatchGetFiles.FileExtensions);
                    searchFunction.FileName = searchQ;
                }

                searchFunction.FindFolders = false;
                searchFunction.StartFolderID = folder.ID;
                searchFunction.Recursive = BatchGetFiles.Recursive;
                var ids = new List<Tuple<int, EdmObjectType>>();

                var searchResult = searchFunction.GetFirstResult();
                while (searchResult != null)
                {
                    ids.Add(new Tuple<int, EdmObjectType>(searchResult.ID, searchResult.ObjectType));

                    searchResult = searchFunction.GetNextResult();
                }

                var ppoObjectsLst = new List<EdmObjectInfo>();
                ids.ForEach(x => ppoObjectsLst.Add(new EdmObjectInfo() { moObjectID = x.Item1 , meType = x.Item2 }));

                var vault9 = vault as IEdmVault9;
                EdmObjectInfo[] ppoObjects = ppoObjectsLst.ToArray();
                vault9.GetObjects(ref ppoObjects);


                return ppoObjects.ToList().Select(x => x.mpoObject as IEdmFile5).ToArray();

            }

        
        }

      

        public static object[] GetAllVariablesValuesFromFolder(this IEdmFolder5 folder, string[] variableNames)
        {
            throw new NotImplementedException();

        }

        /// <summary>
        /// Returns the file type, whether the file is a SOLIDWORKS document or a generic one.
        /// </summary>
        /// <param name="file">File</param>
        /// <returns></returns>
        public static swPDMDocumentType_e GetFileType(this IEdmFile5 file)
        {
            if (file.IsPart())
                return swPDMDocumentType_e.Part;

            if (file.IsDrawing())
                return swPDMDocumentType_e.Drawing;

            if (file.IsAssembly())
                return swPDMDocumentType_e.Assembly;

            return swPDMDocumentType_e.Generic;
        }

        /// <summary>
        /// Gets all subfolders in the given folder
        /// </summary>
        /// <param name="folder"></param>
        /// <returns>Array of subfolders IEdmFolder5[]</returns>
        public static IEdmFolder5[] GetFolders(this IEdmFolder5 folder)
        {
            var foldersList = new List<IEdmFolder5>();
            var position = folder.GetFirstSubFolderPosition();

            while (!position.IsNull)
            {
                foldersList.Add(folder.GetNextSubFolder(position));
            }

            return foldersList.ToArray();
        }


        /// <summary>
        /// Attempts to get <see cref="IEdmFile5"/> from path. This method swallows any exceptions if the file does not exist, and returns null instead.
        /// </summary>
        /// <param name="vault">The vault.</param>
        /// <param name="filePath">The file path.</param>
        /// <param name="folder">The folder.</param>
        /// <returns><see cref="IEdmFile5"/></returns>
        /// <exception cref="System.ArgumentNullException">vault</exception>
        /// <exception cref="System.ArgumentException"></exception>
        public static IEdmFile5 TryGetFileFromPath(this IEdmVault5 vault, string filePath, out IEdmFolder5 folder) 
        {

            var ret = default(IEdmFile5);
            folder = default(IEdmFolder5);

            if (vault == null)
                throw new ArgumentNullException(nameof(vault));

            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException($"{nameof(filePath)} cannot be null or white space.");

            try
            {
                ret = vault.GetFileFromPath(filePath, out folder);
            }
            catch (Exception)
            {

            }


            return ret;
        }



        /// <summary>
        /// Attempts to get <see cref="IEdmFolder5"/> from path. This method swallows any exceptions if the folder does not exist, and returns null instead.
        /// </summary>
        /// <param name="vault">The vault.</param>
        /// <param name="folderPath">The folder path.</param>
        /// <returns><see cref="IEdmFolder5"/></returns>
        /// <exception cref="System.ArgumentNullException">vault</exception>
        /// <exception cref="System.ArgumentException"></exception>
        public static IEdmFolder5 TryGetFolderFromPath(this IEdmVault5 vault, string folderPath)
        {

            var ret = default(IEdmFolder5);
      
            if (vault == null)
                throw new ArgumentNullException(nameof(vault));

            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException($"{nameof(folderPath)} cannot be null or white space.");

            try
            {
                ret = vault.GetFolderFromPath(folderPath);
            }
            catch (Exception)
            {

            }


            return ret;
        }



        /// <summary>
        /// Returns the ID of the parent folder.
        /// </summary>
        /// <param name="file">File</param>
        /// <returns>ID of the parent folder.</returns>
        public static int GetParentFolderID(this IEdmFile5 file)
        {
            var pos = file.GetFirstFolderPosition();

            

            if (pos.IsNull)
                return 0;

            var folder = file.GetNextFolder(pos);
            if (folder == null)
                return 0;
            
            return folder.ID;
        }

        /// <summary>
        /// Returns an enum telling you the check-in state of the file.
        /// </summary>
        /// <param name="file"File</param>
        /// <returns><see cref="CheckoutAction"/></returns>
        public static CheckoutAction GetRequiredCheckOutAction(this IEdmFile5 file)
        {
            var vault = file.Vault as IEdmVault7;

            if (vault.IsLoggedIn == false)
                throw new Exception($"You are not logged into {vault.Name}");

            var userMgr = vault.CreateUtility(EdmUtility.EdmUtil_UserMgr) as IEdmUserMgr5;
            if (file.IsLocked)
            {
                if (file.LockedByUserID == userMgr.GetLoggedInUser().ID)
                    return CheckoutAction.DoNothingFileCheckedOutByMe;
                else
                    return CheckoutAction.CheckedOutBySomeoneElse;
            }
            else
                return CheckoutAction.FileCheckedInCanBeCheckedOut;
        }

        /// <summary>
        /// Returns an array of all variable names.
        /// </summary>
        /// <param name="vault"></param>
        /// <returns></returns>
        public static string[] GetVariableNames(this IEdmVault5 vault)
        {
            var varManager = (vault as IEdmVault7).CreateUtility(EdmUtility.EdmUtil_VariableMgr) as IEdmVariableMgr5;
            var position = varManager.GetFirstVariablePosition();
            var varList = new List<string>();

            while (!position.IsNull)
            {
                var nextVar = varManager.GetNextVariable(position);
                varList.Add(nextVar.Name);
            }

            return varList.ToArray();
        }

        /// <summary>
        /// Returns the different states of the local copy of the file.
        /// </summary>
        /// <param name="file">File</param>
        /// <returns></returns>
        public static LocalCopyState HasLocalCopy(this IEdmFile5 file)
        {
            bool pbLocalOverwrittenVersionObsolete = false;
            var currentVersion = file.CurrentVersion;
            var localVersion = (file as IEdmFile12).GetLocalVersionNo2(file.GetParentFolderID(), out pbLocalOverwrittenVersionObsolete);
            if (localVersion == -1)
                return LocalCopyState.ThereIsNoLocalCopy;
            else
            {
                if (pbLocalOverwrittenVersionObsolete)
                    return LocalCopyState.LocalCopyObseleteServerHasNewOne;
                else
                {
                    if (localVersion == currentVersion)
                        return LocalCopyState.LocalAndServerCopySynced;
                    else
                        return LocalCopyState.LocalCopyIsBehindServer;
                }
            }
        }

        /// <summary>
        /// Returns whether the document is a SOLIDWORKS assembly or not.
        /// </summary>
        /// <param name="file">File.</param>
        /// <returns>File</returns>
        public static bool IsAssembly(this IEdmFile5 file)
        {
            return file.Name.ToLower().EndsWith(".sldasm");
        }

        /// <summary>
        /// Returns whether the file is checked by the currently logged in user or not.
        /// </summary>
        /// <param name="file">File</param>
        /// <returns>True if checked out by the currently logged in user, false if not.</returns>
        public static bool IsCheckedOutByMe(this IEdmFile5 file)
        {
            var vault = file.Vault as IEdmVault7;

            if (vault.IsLoggedIn == false)
                throw new Exception($"You are not logged into {vault.Name}");

            var userMgr = vault.CreateUtility(EdmUtility.EdmUtil_UserMgr) as IEdmUserMgr5;
            if (file.LockedByUserID == userMgr.GetLoggedInUser().ID)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Returns whether the document is a SOLIDWORKS drawing or not.
        /// </summary>
        /// <param name="file">File.</param>
        /// <returns>File</returns>
        public static bool IsDrawing(this IEdmFile5 file)
        {
            return file.Name.ToLower().EndsWith(".slddrw");
        }

        /// <summary>
        /// Returns whether the document is a SOLIDWORKS part or not.
        /// </summary>
        /// <param name="file">File.</param>
        /// <returns>File</returns>
        public static bool IsPart(this IEdmFile5 file)
        {
            return file.Name.ToLower().EndsWith(".sldprt");
        }

        /// <summary>
        /// Rollbacks the file to a previous version.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="version">Previous version.</param>
        /// <returns></returns>
        public static void RollbackToVersion(this IEdmFile5 file, int version, int handle = 0)
        {
            if (file.CurrentVersion <= 1)
                throw new Exception("File is in version 1");

            var checkoutAction = file.GetRequiredCheckOutAction();
            var folderID = file.GetParentFolderID();
            var fileInfo = new FileInfo(file.GetLocalPath(folderID));
            var directory = fileInfo.Directory.FullName;
            switch (checkoutAction)
            {
                case CheckoutAction.DoNothingFileCheckedOutByMe:
                    break;

                case CheckoutAction.FileCheckedInCanBeCheckedOut:
                    file.LockFile(folderID, handle, (int)EdmLockFlag.EdmLock_Simple);
                    break;

                case CheckoutAction.CheckedOutBySomeoneElse:
                    throw new Exception("The file is checked out by someone else.");
                default:
                    break;
            }

            file.GetFileCopy(handle, version, folderID, (int)EdmGetFlag.EdmGet_Simple, string.Empty);

            file.UnlockFile(handle, $"Rolled back to version {version}.");
        }

        /// <summary>
        /// Attempts to get a strongly typed object from the vault using its ID. Returns nothing or null if it fails.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="vault">Vault object.</param>
        /// <param name="objectID">ID of the object.</param>
        /// <returns></returns>
        public static T TryGetObject<T>(this IEdmVault5 vault, int objectID)
        {
            if (vault == null)
                throw new ArgumentNullException("vault");

            try
            {
                var typeName = typeof(T).Name;
                var typeNameSanitized = typeName.Replace("IEdm", "").TrimEnd(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
                var types = Enum.GetValues(typeof(EdmObjectType));
                var foundtype = default(EdmObjectType);
                var matchingEndingStr = $"_{typeNameSanitized.ToLower()}".Trim();
                foreach (var type in types)
                {
                    var swtype = (EdmObjectType)type;
                    var swtypeStr = swtype.ToString().ToLower().Trim();

                    if (swtypeStr.EndsWith(matchingEndingStr, StringComparison.InvariantCulture))
                    {
                        foundtype = swtype;
                        break;
                    }
                }
                var ret = (T)vault.GetObject(foundtype, objectID);

                return ret;
            }
            catch (Exception e)
            {
                return default(T);
            }
        }

        #endregion
    }
}