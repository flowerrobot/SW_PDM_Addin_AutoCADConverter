﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using Newtonsoft.Json;
using EPDM.Interop.epdm;

namespace AutoCADConversion
{


    [Guid("51B90262-FEAE-4F6B-9FD8-15625DF9595E"), ComVisible(true)]
    public class PDMAddin : IEdmAddIn5
    {
     internal static   NLog.Logger log = NLog.LogManager.GetLogger("AutoCAD Converter");
        public static IEdmVault20 Vault { get; private set; }
        public static IEdmCmdMgr5 CmdMgr { get; private set; }
        public static EdmAddInInfo Info { get; private set; }
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            Vault = (IEdmVault20)poVault;
            CmdMgr = poCmdMgr;
            Info = poInfo;

            //Specify information to display in the add-in's Properties dialog box
            poInfo.mbsAddInName = "IAC Debug Add-in2";
            poInfo.mbsCompany = "IAC";
            poInfo.mbsDescription = "This is a debug testing for IAC addin";
            poInfo.mlAddInVersion = 10;

            //Specify the minimum required version of SolidWorks PDM Professional
            poInfo.mlRequiredVersionMajor = 11;
            poInfo.mlRequiredVersionMinor = 1;

            // Register a menu command
            // poCmdMgr.AddCmd(1, "C# Add-in", (int)EdmMenuFlags.EdmMenu_ShowInMenuBarTools);

            poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardButton);


            //Register this add-in as a task add-in
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun);
            //Register this add-in to be called when selected as a task in the Administration tool
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup);
            //Register this add-in to be called when the task is launched on the client computer
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskLaunch);
            //Register this add-in to provide extra details in the Details dialog box in the task list in the Administration tool
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskDetails);
            //Register this add-in to be called when the launch dialog box is closed
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskLaunchButton);
            //Register this add-in to be called when the set-up wizard is closed
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetupButton);

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="poCmd"></param>
        /// <param name="ppoData"> is EdmCmdData()</param>
        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {
                PauseToAttachProcess(poCmd.meCmdType.ToString());

                // Handle the menu command
                switch (poCmd.meCmdType)
                {
                    case EdmCmdType.EdmCmd_Menu:
                        switch (poCmd.mlCmdID)
                        {
                            case 1:
                                // ClickToRunTask(poCmd);
                                break;
                        }
                        break;
                    case EdmCmdType.EdmCmd_CardButton:
                        break;
                    case EdmCmdType.EdmCmd_TaskRun:
                        OnTaskRun(ref poCmd, ref ppoData);
                        break;
                    case EdmCmdType.EdmCmd_TaskSetup:
                        OnTaskSetup(ref poCmd, ref ppoData);
                        break;
                    case EdmCmdType.EdmCmd_TaskDetails:
                        OnTaskDetails(ref poCmd, ref ppoData);
                        break;
                    case EdmCmdType.EdmCmd_TaskLaunch:
                    case EdmCmdType.EdmCmd_TaskLaunchButton:
                    case EdmCmdType.EdmCmd_TaskSetupButton:
                        OnTaskSetupButton(ref poCmd, ref ppoData);
                        break;
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                log.Error(ex);
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                log.Error(ex);
                MessageBox.Show(ex.Message);
            }
        }

        //public void ClickToRunTask(EdmCmd poCmd)
        //{
        //    var vault = (IEdmVault19)poCmd.mpoVault;
        //    IEdmTaskMgr TaskMgr = (IEdmTaskMgr)vault.CreateUtility(EdmUtility.EdmUtil_TaskMgr);
        //    //var Tasks = 

        //    foreach(EdmTaskInfo task in TaskMgr.GetTasks())
        //    {
        //        if(task.mbsTaskName == "MyTask")
        //        {
        //            TaskMgr.RunTask(task, null, 0);
        //        }
        //    }

        //}

        private void OnTaskRun(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            //Get the task instance interface
            IEdmTaskInstance inst = poCmd.mpoExtra as IEdmTaskInstance;
            if (inst == null)
                return;
            try
            {
                var tas = new AutoCADTaskAddin(poCmd, ppoData);
                tas.runTask();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                log.Error(ex);
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, ex.ErrorCode, "The test task failed!");
            }
            catch (Exception ex)
            {
                log.Error(ex);
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, 0, "", "The test task failed!");
            }
        }

        AutoCADTaskSettings taskSettings;
        private void OnTaskSetup(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {

                //Get the property interface used to access the framework
                IEdmTaskProperties props = poCmd.mpoExtra as IEdmTaskProperties;
                //Get the settings for exist task               
                try
                {
                    taskSettings = JsonConvert.DeserializeObject<AutoCADTaskSettings>(props.GetValEx(AutoCADTaskSettings.Acadtask_Settings) as string);
                }
                catch
                {
                    taskSettings = new AutoCADTaskSettings();
                }

                //Set the property flag that says you want a
                //menu item for the user to launch the task
                //and a flag to support scheduling
                props.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsInitExec + (int)EdmTaskFlag.EdmTask_SupportsDetails + (int)EdmTaskFlag.EdmTask_SupportsChangeState + (int)EdmTaskFlag.EdmTask_SupportsInitForm;



                EdmTaskSetupPage[] setupPages = new EdmTaskSetupPage[2];
                setupPages[0].mbsPageName = "Menus Command";
                setupPages[0].mbsPageName = "TitleBlocks";
                //pages[0].mlPageHwnd = SetupPageObj.Handle.ToInt32();
                //pages[0].mpoPageImpl = SetupPageObj;
                props.SetSetupPages(setupPages);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                System.IO.File.OpenWrite(@"C:\temp\errorfile.text");
                // MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                System.IO.File.OpenWrite(@"C:\temp\errorfile.text");
                //MessageBox.Show(ex.Message);
            }
        }
        private void OnTaskSetupButton(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {
                //Custom setup page, SetupPageObj, is created
                //in method Class1::OnTaskSetup; SetupPage::StoreData 
                //saves the contents of the list box to poCmd.mpoExtra 
                //in the IEdmTaskProperties interface
                if (poCmd.mbsComment == "OK" && (taskSettings != null))
                {
                    IEdmTaskProperties props = poCmd.mpoExtra as IEdmTaskProperties;
                    //Pull settings from windows


                    //Set up the menu commands to launch this task
                    if (taskSettings.CreateMenu)
                    {
                        EdmTaskMenuCmd[] cmds = new EdmTaskMenuCmd[1];
                        cmds[0].mbsMenuString = taskSettings.MenuName;
                        cmds[0].mbsStatusBarHelp = taskSettings.MenuDescription;
                        cmds[0].mlCmdID = 2;
                        cmds[0].mlEdmMenuFlags = (int)EdmMenuFlags.EdmMenu_MustHaveSelection;
                        props.SetMenuCmds(cmds);
                    }

                    props.SetValEx(AutoCADTaskSettings.Acadtask_Settings, JsonConvert.SerializeObject(taskSettings));
                }
                taskSettings = null;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void OnTaskDetails(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {
                IEdmTaskInstance TaskInstance = (IEdmTaskInstance)poCmd.mpoExtra;
                if ((TaskInstance != null))
                {
                    // SetupPageObj;//= new SetupPage((IEdmVault7)poCmd.mpoVault, TaskInstance);
                    return;
                    //Force immediate creation of the control
                    //and its handle
                    //SetupPageObj.CreateControl();
                    // SetupPageObj.LoadData(poCmd);
                    // SetupPageObj.DisableControls();
                    poCmd.mbsComment = "State Age Details";
                    //  poCmd.mlParentWnd = SetupPageObj.Handle.ToInt32();
                }

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void PauseToAttachProcess(string callbackType)
        {
            try
            {
                //If the debugger isn't already attached to a
                //process, 
                if (!Debugger.IsAttached)
                {
                    //Launch the debug dialog
                    //Debugger.Launch()
                    //or
                    //use a MsgBox dialog to pause execution
                    //and allow the user time to attach it
                    MessageBox.Show("Attach debugger to process \"" + Process.GetCurrentProcess().ProcessName + "\" for callback \"" + callbackType + "\" before clicking OK.");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
