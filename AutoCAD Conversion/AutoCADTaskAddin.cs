using EPDM.Interop.epdm;
using Newtonsoft.Json;
using Syncfusion.Pdf;
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
        AutoCADTaskSettings settings;
        public AutoCADTaskAddin(EdmCmd poCmd, EdmCmdData[] ppoData, AutoCADTaskSettings taskSettings)
        {
            //Get the task instance interface
            inst = poCmd.mpoExtra as IEdmTaskInstance;
            data = ppoData;
            settings = taskSettings;
            if (inst == null)
                throw new Exception("Task values incorrect");

            vault = (IEdmVault20)poCmd.mpoVault;

            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("MTcwMjQwQDMxMzcyZTMzMmUzMGhycFlDaldXNDVZeWxhdnFwckswQnRhMHVwclp2OWNrUEltNHczb21ENDQ9;MTcwMjQxQDMxMzcyZTMzMmUzMG4xOFQ1dnBDR1oxalUvazM5UmlTRkdUelJRcHkweURnVERXRXRabnpaZVE9");
        }

        public void runTask()//EdmCmdData[] data
        {
            List<string> file2Delete = new List<string>();
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


                file2Delete.Add(dataName);
                file2Delete.Add(scriptPath);
                file2Delete.Add(lispPath);


                //collect data               
                foreach (EdmCmdData item in data)
                {
                    if (item.mlLongData1 == (long)EdmObjectType.EdmObject_File)
                    {
                        var fileParm = new AutoCADParameters();
                        fileParm.FileId = item.mlObjectID1;
                        fileParm.FolderId = item.mlObjectID2;
                        fileParm.FileName = item.mbsStrData1;

                        ///Custom code to output path
                        var path = Path.GetDirectoryName(fileParm.FileName).Split('\\');
                        try
                        {
                            if (path.Length > 3 && path[2].Equals("30 Projects"))
                            {
                                fileParm.OutputPath = Path.Combine(path[0] + @"\", path[1], path[2], path[3], "03 PDFs", Path.GetFileNameWithoutExtension(fileParm.FileName) + ".pdf");
                                fileParm.OutputPathDXF = Path.Combine(path[0] + @"\", path[1], path[2], path[3], "04 DXFs", Path.GetFileNameWithoutExtension(fileParm.FileName) + ".dxf");
                            }
                            else
                            {
                                fileParm.OutputPath = Path.ChangeExtension(fileParm.FileName, ".pdf");
                                fileParm.OutputPathDXF = Path.ChangeExtension(fileParm.FileName, ".dxf");
                            }
                        }
                        catch
                        {
                            fileParm.OutputPath = Path.ChangeExtension(fileParm.FileName, ".pdf");
                            fileParm.OutputPathDXF = Path.ChangeExtension(fileParm.FileName, ".dxf");
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
                            var allFiles = Directory.GetFiles(Path.GetDirectoryName(OutputfilePath), Path.GetFileNameWithoutExtension(OutputfilePath) + "*");
                            if (allFiles.Count() > 1)
                            {
                                List<string> files2Merge = new List<string>();
                                foreach (string f in allFiles)
                                {
                                    if (f.Equals(OutputfilePath, StringComparison.OrdinalIgnoreCase))
                                        file2Delete.Add(OutputfilePath);
                                    else if (Path.GetExtension(f).Equals(".dxf", StringComparison.OrdinalIgnoreCase))
                                        file2Delete.Add(Path.ChangeExtension(OutputfilePath, ".dxf"));
                                    else
                                    {
                                        files2Merge.Add(f);
                                        file2Delete.Add(f);
                                    }
                                }
                                if (files2Merge.Count > 0)
                                {
                                    files2Merge.Insert(0, OutputfilePath);
                                    PdfDocument finalDoc = new PdfDocument();
                                    PdfDocument.Merge(finalDoc, files2Merge.ToArray());
                                    File.Delete(OutputfilePath);
                                    finalDoc.Save(OutputfilePath);
                                    finalDoc.Close(true);
                                }
                            }
                            else
                            { 
                                file2Delete.Add(OutputfilePath);
                                file2Delete.Add(Path.ChangeExtension(OutputfilePath, ".dxf")); }

                            IEdmFile5 filePDF = CopyFileIntoVault(OutputfilePath, fileParm.OutputPath);
                            IEdmFile5 fileDXF = CopyFileIntoVault(Path.ChangeExtension(OutputfilePath, ".dxf"), fileParm.OutputPathDXF);


                            //If pdf is now in the vault
                            if (filePDF != null)
                            {

                                //Update variables, first get a copy of the files from the source
                                IEdmBatchListing4 var = (IEdmBatchListing4)vault.CreateUtility(EdmUtility.EdmUtil_BatchList);
                                var.AddFileCfg(fileParm.FileName, default(DateTime), 0);
                                var colsetNames = var.GetColumnSetNames();

                                //string atts = "";
                                //foreach (VariableMapperViewModel va in settings.Variables)
                                //{
                                //    if (va.MapVariable)
                                //        atts += "\n" + va.SourceVariable.Name;
                                //}


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
                                        bUp.SetVar(filePDF.ID, cols[i].mlVariableID, ((string[])getFile.moColumnData)[i], "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                        bUp.SetVar(fileDXF.ID, cols[i].mlVariableID, ((string[])getFile.moColumnData)[i], "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                    }
                                }
                                //Set SourcedState
                                bUp.SetVar(filePDF.ID, 171, "Released", "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                bUp.SetVar(fileDXF.ID, 171, "Released", "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                //Set AutomaticallyGenerated
                                bUp.SetVar(filePDF.ID, 158, 1, "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                bUp.SetVar(fileDXF.ID, 158, 1, "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                //foreach (EdmListFile2 getFile in files)
                                //{                                    
                                //    for (int i = 0; i < cols.Length; i++)
                                //    {
                                //        VariableMapperViewModel va = settings.Variables.FirstOrDefault(t=> t.SourceVariable.Id == ((EdmListCol[])cols)[i].mlVariableID);
                                //        bUp.SetVar(Outfile.ID, va.DestinationVariable.Id , ((string[])getFile.moColumnData)[i], "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                //    }
                                //}

                                ////copy the static text
                                //foreach (VariableMapperViewModel va in settings.Variables)
                                //{
                                //    if (!va.MapVariable)
                                //        bUp.SetVar(Outfile.ID, va.DestinationVariable.Id, va.Value, "", (int)EdmBatchFlags.EdmBatch_AllConfigs);
                                //}

                                if (0 != bUp.CommitUpdate(out EdmBatchError2[] err))
                                {
                                    PDMAddin.log.Error("Variable errors");
                                }
                                //Check file in.
                                //IEdmFolder5 fldPDF = vault.GetFolderFromPath(Path.GetDirectoryName(fileParm.OutputPath));
                                // IEdmFolder5 fldDXF = vault.GetFolderFromPath(Path.GetDirectoryName(fileParm.OutputPathDXF));

                                filePDF.UnlockFile(0, "Automaticly created", (int)EdmUnlockFlag.EdmUnlock_IgnoreReferences);
                                fileDXF.UnlockFile(0, "Automaticly created", (int)EdmUnlockFlag.EdmUnlock_IgnoreReferences);
                            }
                        }
                    }
                }

                ///comlete

                inst.SetProgressPos(100, "finished");
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneOK);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                PDMAddin.log.Error(ex);
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, ex.ErrorCode, "The task failed!");
            }
            catch (Exception ex)
            {
                PDMAddin.log.Error(ex);
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, 0, "", "The task failed!");
            }
            finally
            {
                foreach (string f in file2Delete)
                {
                    if (File.Exists(f))
                        File.Delete(f);
                }
            }
        }
        public IEdmFile5 CopyFileIntoVault(string path, string Destination)
        {
            IEdmFile5 Outfile = vault.GetFileFromPath(Destination, out IEdmFolder5 fld);
            if (Outfile != null)
            {
                if (!Outfile.IsLocked)
                {
                    Outfile.LockFile(fld.ID, 0, 0);
                }
                System.IO.File.Copy(path, Destination, true);
                return Outfile;
            }
            else
            {
                fld = vault.GetFolderFromPath(Path.GetDirectoryName(Destination));
                if (fld != null)
                {
                    int id = ((IEdmFolder8)fld).AddFile2(0, path, out int err, "", (int)EdmAddFlag.EdmAdd_UniqueVarDelayCheck);
                    return (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, id);
                }
                return null;
            }
        }
    }
}
