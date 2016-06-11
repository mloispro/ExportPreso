using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.IO;

namespace ExportPreso
{
    public struct SlideInfo
    {
        public int Id;
        public string path;
    }
    class PowerPointEx
    {
        
        //private IntPtr _hWndPowerPoint;
        //private IntPtr _hwndHost;

        public static Microsoft.Office.Interop.PowerPoint.Application Start()
        {
            Microsoft.Office.Interop.PowerPoint.Application powerpointApp = null;
            // Startup PowerPoint using the Office interop library. We're using version 12 here, since the customer uses Office 2007
            if (null == powerpointApp)
            {
                // First, check if POWERPNT.EXE is a running application. Can be an unwanted left over or a manually started PowerPoint instance
                // Could (probably) also be another application using PowerPoint interops, in which case we're stealing their instance. 
                // Might interfere with that other application, but this control is not intended for side by side use

                //Get reference to Excel.Application from the ROT.
                try
                {
                    powerpointApp = (Microsoft.Office.Interop.PowerPoint.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Powerpoint.Application");
                    if (powerpointApp != null)
                    {
                        powerpointApp.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized;
                    }else
                    {
                        powerpointApp = new Microsoft.Office.Interop.PowerPoint.Application();
                    }
                }
                catch (COMException)
                {
                    // If no instance was running, this exception is thrown. Should be the case under normal circumstances
                    // Also sometimes (maybe only during development), an exception is thrown, mentioning open dialogs.
                    powerpointApp = null;

                }
                catch (Exception ex)
                {
                    // Something else went wrong
                   // Logger.LogException(ex);
                }

                if (powerpointApp==null )
                {
                    

                    // Now, launch a new instance
                    try
                    {
                        powerpointApp = new Microsoft.Office.Interop.PowerPoint.Application();
                   
                    }
                    catch (COMException ex)
                    {
                        //Logger.LogException(ex);
                    }
                }

                if (null == powerpointApp)
                {
                    // We couldn't start. Maybe PowerPoint isn't installed?
                    //Logger.LogWarning(@"Could not launch PowerPoint 2007. Is it installed?");
                    //return false;
                }
                
            }
            return powerpointApp;
        }

        public static bool Close(Microsoft.Office.Interop.PowerPoint.Application powerpointApp)
        {
            bool closed = false;
            // If no PowerPoint instance was running, check for a running process. It could be a manual PowerPoint was started, which we can't connect to
            try
            {
                powerpointApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(powerpointApp);
                Process[] pros = Process.GetProcesses();
                for (int i = 0; i < pros.Count(); i++)
                {
                    if (pros[i].ProcessName.ToLower().Contains("powerpnt"))
                    {
                        pros[i].Kill();
                        closed = true;
                    }
                }
                if (closed)
                    return true;

                Process[] powerPointsRunning = Process.GetProcessesByName("POWERPNT.EXE");
                if (powerPointsRunning.Length > 0) // Yes, it is
                {
                    foreach (var powerpointRunning in powerPointsRunning)
                    {
                        powerpointRunning.Kill(); // Just kill it
                        closed = true;
                    }
                    Thread.Sleep(500); // Half a second should do it. Open dialogs and changes are ignored. Tough luck
                }
                if (closed)
                    return true;
            }
            catch (Exception ex)
            {
                // For good measure, catch exceptions and log them
                //Logger.LogException(ex);
            }
            Process[] powerPointsRunning2 = Process.GetProcessesByName("POWERPNT.EXE");
            if (powerPointsRunning2.Length > 0)
            {
                // We were unsuccessful in closing PowerPoint. Report and return false
                //Logger.LogWarning(@"Unable to close down running instance of PowerPoint.");
                //return false;
            }
            return closed;
        }
        public static Microsoft.Office.Interop.PowerPoint.Application ConvertToPPTX(List<string> sourcePpts, string tempDir)
        {
            //Close(powerpointApp);
            Microsoft.Office.Interop.PowerPoint.Application newPowerpointApp = Start();

            string targetPptx = "";
            //List<SlideInfo> slideInfos = new List<SlideInfo>();
            foreach (var sourcePpt in sourcePpts)
            {
                
                string pptFile = Path.GetFileName(sourcePpt);
                targetPptx = Path.Combine(tempDir, pptFile);
            
                if (!sourcePpt.Contains(".pptx"))
                    targetPptx += "x";

                if (!File.Exists(targetPptx))
                {
                    //Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(preso, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                    //Microsoft.Office.Interop.PowerPoint.Presentation ppt = newPowerpointApp.Presentations.Open(sourcePpt, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoTrue);
                    Microsoft.Office.Interop.PowerPoint.Presentation prez = newPowerpointApp.Presentations.Open(sourcePpt, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                    prez.SaveAs(targetPptx, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault);
                    prez.Close();
                }
            }
            Close(newPowerpointApp);
            Microsoft.Office.Interop.PowerPoint.Application returnPowerpointApp = Start();
           // Microsoft.Office.Interop.PowerPoint.Presentation pptx = returnPowerpointApp.Presentations.Open(targetPptx, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoTrue);
            return returnPowerpointApp;

        }
        private void _powerpointApp_SlideShowNextSlide(SlideShowWindow Wn)
        {
            //throw new NotImplementedException();
        }

        private void powerpointApp_SlideShowEnd(Presentation Pres)
        {
            //throw new NotImplementedException();
        }

        private void powerpointApp_SlideShowBegin(SlideShowWindow Wn)
        {
            //throw new NotImplementedException();
        }
    }
}
