using EPDM.Interop.epdm;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCADConversion
{
    public class AutoCADTaskAddin
    {
        const string CorePath = @"C:\Program Files\Autodesk\AutoCAD 2018\accoreconsole.exe";

        IEdmVault20 vault;
        IEdmTaskInstance inst;
        EdmCmdData[] data;
        public AutoCADTaskAddin(EdmCmd poCmd, EdmCmdData[] ppoData)
        {
            //Get the task instance interface
            inst = poCmd.mpoExtra as IEdmTaskInstance;
            data = ppoData;
            if (inst == null)
                throw new Exception("Task values incorrect");

            vault = (IEdmVault20)poCmd.mpoVault;
        }

        public void runTask()//EdmCmdData[] data
        {
            try
            {
                //Keep the Task List status up to date
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_Running);
                inst.SetProgressRange(100, 0, "Starting");

                string tempDir = Path.GetTempPath() + @"AutoCadTaskConverter\";
                string dataName = tempDir + @"proccessFile.txt";
                string scriptPath = tempDir + @"proccessScript.scr";
                string lispPath = tempDir + @"proccessLisp.lsp";
                if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);

                //collect data               
                foreach (EdmCmdData item in data)
                {
                    if (item.mlLongData1 == (long)EdmObjectType.EdmObject_File)
                    {
                        var fileParm = new AutoCADParameters();
                        fileParm.FileId = item.mlObjectID1;
                        fileParm.FolderId = item.mlObjectID2;
                        fileParm.FileName = item.mbsStrData1;

                        var path = Path.GetDirectoryName(fileParm.FileName).Split('\\');
                        try
                        {
                            if (path.Length > 3 && path[2].Equals("30 Projects"))
                            {
                                fileParm.OutputPath = Path.Combine(path[0] + @"\", path[1], path[2], path[3], "03 PDFs", Path.GetFileNameWithoutExtension(fileParm.FileName) + ".pdf");
                            }
                            else
                            {
                                fileParm.OutputPath = Path.ChangeExtension(fileParm.FileName, ".pdf");
                            }
                        }
                        catch
                        {
                            fileParm.OutputPath = Path.ChangeExtension(fileParm.FileName ,".pdf");
                        }


                        var CADFile = vault.GetFileFromPath(fileParm.FileName, out IEdmFolder5 Cadfld);
                        if (CADFile == null) { break; }

                        string OutputfilePath = tempDir + Path.GetFileNameWithoutExtension(fileParm.FileName) + ".pdf";

                        if (System.IO.File.Exists(OutputfilePath))
                            File.Delete(OutputfilePath);


                        inst.SetProgressPos(10, "Writting report");
                        ///Write json file ensure file is local
                        JsonSerializer serializer = new JsonSerializer();
                        using (StreamWriter sw = new StreamWriter(dataName, false))
                        {
                            using (JsonWriter writer = new JsonTextWriter(sw))
                            {
                                serializer.Serialize(writer, fileParm);
                            }
                        }
                        inst.SetProgressPos(15, "Launching application");


                        ///Launch autocad.
                        File.WriteAllText(scriptPath, string.Format(AutoCADConversion.Properties.Resources.AcadScript, lispPath.Replace(@"\", @"\\")));
                        File.WriteAllText(lispPath, string.Format(AutoCADConversion.Properties.Resources.ACADLisp, @"""" + OutputfilePath.Replace(@"\", @"\\") + @""""));

                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo
                        {
                            CreateNoWindow = false,
                            UseShellExecute = true,
                            FileName = CorePath,
                            WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal,
                            Arguments = $"/i \"{fileParm.FileName}\" /s \"{scriptPath}\""
                        };

                        using (System.Diagnostics.Process exeProccess = System.Diagnostics.Process.Start(startInfo))
                        {
                            inst.SetProgressPos(20, "Launching application");
                            exeProccess?.WaitForExit();
                        }

                        //If Pdf has generated, then copy the file to the desitnation output
                        if (System.IO.File.Exists(OutputfilePath))
                        {
                            var Outfile = vault.GetFileFromPath(fileParm.OutputPath, out IEdmFolder5 fld);
                            int id = 0;
                            if (Outfile != null)
                            {
                                if (!Outfile.IsLocked)
                                {
                                    Outfile.LockFile(fld.ID, 0, 0);
                                }
                                System.IO.File.Copy(OutputfilePath, fileParm.OutputPath, true);
                                id = Outfile.ID;
                            }
                            else
                            {
                                fld = vault.GetFolderFromPath(Path.GetDirectoryName(fileParm.OutputPath));
                                if (fld != null)
                                {
                                    id = ((IEdmFolder8)fld).AddFile2(0, OutputfilePath, out int err, "", (int)EdmAddFlag.EdmAdd_UniqueVarDelayCheck);
                                    Outfile = (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, id);
                                }
                            }

                            //If pdf is now in the vault
                            if (id != 0)
                            {

                                //Update variables, first get a copy of the files from the source
                                IEdmBatchListing4 var = (IEdmBatchListing4)vault.CreateUtility(EdmUtility.EdmUtil_BatchList);
                                var.AddFileCfg(Outfile.ID, default(DateTime), 0);
                                var colsetNames = var.GetColumnSetNames();
                                string atts = "\nDescription\nDescription2\nDescription3\nRevision\nDocument Number\nDrawnBy\nApprovedBy\nProject Name\nProject Number";
                                EdmListCol[] cols = default(EdmListCol[]);
                                EdmListFile2[] files = default(EdmListFile2[]);
                                var.CreateListEx(atts, (int)EdmCreateListExFlags.Edmclef_MayReadFiles, ref cols, null);
                                var.GetFiles2(ref files);

                                //Now copy the files to the variables.
                                IEdmBatchUpdate2 bUp = (IEdmBatchUpdate2)vault.CreateUtility(EdmUtility.EdmUtil_BatchUpdate);
                                foreach (EdmListFile2 getFile in files)
                                {
                                    for (int i = 0; i < cols.Length; i++)
                                    {
                                        bUp.SetVar(Outfile.ID, ((EdmListCol[])cols)[i].mlVariableID, ((string[])getFile.moColumnData)[i], "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                    }
                                }
                                //Check file in.
                                Outfile.UnlockFile(0, "Automaticly created", (int)EdmUnlockFlag.EdmUnlock_IgnoreReferences);
                            }
                        }

                        //System.IO.File.Delete(dataName);
                        //System.IO.File.Delete(scriptPath);
                        //System.IO.File.Delete(lispPath);

                        ///collect pdf \ dxf
                        ///add to vault,
                    }
                }

                ///comlete

                inst.SetProgressPos(100, "finished");
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneOK);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                PDMAddin. log.Error(ex);
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, ex.ErrorCode, "The test task failed!");
            }
            catch (Exception ex)
            {
                PDMAddin. log.Error(ex);
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, 0, "", "The test task failed!");
            }

        }
    }
}
