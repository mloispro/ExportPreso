using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;

namespace ExportPreso
{
    class PowerPointEx
    {
        Microsoft.Office.Interop.PowerPoint.Application _powerpointApp;
        private IntPtr _hWndPowerPoint;
        private IntPtr _hwndHost;

        public bool EnsurePowerPoint()
        {
            // Startup PowerPoint using the Office interop library. We're using version 12 here, since the customer uses Office 2007
            if (null == _powerpointApp)
            {
                // First, check if POWERPNT.EXE is a running application. Can be an unwanted left over or a manually started PowerPoint instance
                // Could (probably) also be another application using PowerPoint interops, in which case we're stealing their instance. 
                // Might interfere with that other application, but this control is not intended for side by side use

                //Get reference to Excel.Application from the ROT.
                try
                {
                    _powerpointApp = (Microsoft.Office.Interop.PowerPoint.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Powerpoint.Application");
                    if (_powerpointApp != null)
                    {
                        _powerpointApp.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized;
                    }
                }
                catch (COMException)
                {
                    // If no instance was running, this exception is thrown. Should be the case under normal circumstances
                    // Also sometimes (maybe only during development), an exception is thrown, mentioning open dialogs.
                    _powerpointApp = null;

                }
                catch (Exception ex)
                {
                    // Something else went wrong
                   // Logger.LogException(ex);
                }

                if (null == _powerpointApp)
                {
                    // If no PowerPoint instance was running, check for a running process. It could be a manual PowerPoint was started, which we can't connect to
                    try
                    {
                        Process[] powerPointsRunning = Process.GetProcessesByName("POWERPNT.EXE");
                        if (powerPointsRunning.Length > 0) // Yes, it is
                        {
                            foreach (var powerpointRunning in powerPointsRunning)
                            {
                                powerpointRunning.Kill(); // Just kill it
                            }
                            Thread.Sleep(500); // Half a second should do it. Open dialogs and changes are ignored. Tough luck
                        }
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
                        return false;
                    }

                    // Now, launch a new instance
                    try
                    {
                        _powerpointApp = new Microsoft.Office.Interop.PowerPoint.Application();
                        if (_powerpointApp != null)
                        {
                            // This will make sure the PowerPoint windows are fully created. It causes some flickering, but we can't do much about that. We need the Windows created before we start the slideshow
                            _powerpointApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                            _powerpointApp.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized;
                        }
                    }
                    catch (COMException ex)
                    {
                        //Logger.LogException(ex);
                    }
                }

                if (null == _powerpointApp)
                {
                    // We couldn't start. Maybe PowerPoint isn't installed?
                    //Logger.LogWarning(@"Could not launch PowerPoint 2007. Is it installed?");
                    return false;
                }
                // Set the window state to minimized will make the startup of PowerPoint much less intrusive.
                _powerpointApp.Assistant.On = false; // We don't want an assistant popping up...
                _powerpointApp.SlideShowBegin += powerpointApp_SlideShowBegin;
                _powerpointApp.SlideShowEnd += powerpointApp_SlideShowEnd;
                _powerpointApp.SlideShowNextSlide += _powerpointApp_SlideShowNextSlide;

                _hWndPowerPoint = (IntPtr)_powerpointApp.HWND;
                if (_hwndHost != IntPtr.Zero)
                {
                    //AttachWindow(_hWndPowerPoint, _hwndHost, (IntPtr)Constants.HWND_TOP, Constants.SWP_HideWindow | Constants.SWP_DoNotChangeOwnerZOrder | Constants.SWP_IgnoreZOrder);
                }
            }
            return _hWndPowerPoint != IntPtr.Zero;
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
