using System;
using System.Collections.Generic;
using EnvDTE;

namespace VisualStudioLauncher
{
    class VSProcess
    {
        #region Constants

        public const int MaxInitializationRetries = 40;
        public const int InitializationRetryMillisecondsTimeout = 100;
        public const int MaxInstanceStartRetries = 60;
        public const int InstanceStartRetryTimeout = 100;

        #endregion

        #region Public Properties

        public string ROTName { get; private set; }
        public EnvDTE80.DTE2 DTE2 { get; private set; }
        public bool HasDTE => DTE2 != null;
        public bool IsInitialized { get; private set; }
        public Solution Solution { get; private set; }
        public bool HasOpenSolution => IsInitialized && Solution != null && Solution.IsOpen;

        #endregion

        #region Public Static Methods

        public static void Initialize(VSProcess vsProcess)
        {
            vsProcess?.Initialize();
        }

        public static void Initialize(List<VSProcess> vsProcesses)
        {
            if (vsProcesses != null)
            {
                foreach (var vsProcess in vsProcesses)
                {
                    vsProcess?.Initialize();
                }
            }
        }

        public static VSProcess FromROT(string ROTName, object ROTObject)
        {
            var dteProcess = (EnvDTE80.DTE2)ROTObject;
            if (dteProcess == null)
                return null;

            return new VSProcess(ROTName, dteProcess);
        }

        public static VSProcess StartNewProcess()
        {
            var process = System.Diagnostics.Process.Start("devenv.exe");
            if (process != null && !process.HasExited)
            {
                //Trick: Wait until process initialized, which happens to be when the window title was given
                //Also, VS seems to give "Visual Studio" once project was loaded so wait for that too..
                int tries = 0;
                while (tries <= MaxInstanceStartRetries &&
                       !process.HasExited && (string.IsNullOrEmpty(process.MainWindowTitle) || !process.MainWindowTitle.EndsWith("Visual Studio")))
                {
                    tries++;
                    System.Threading.Thread.Sleep(InstanceStartRetryTimeout);
                    process.Refresh();
                }

                // Magic value for sleeping before we can hook into the DTE
                System.Threading.Thread.Sleep(1000);
                return VSFinder.GetInstanceByProcessId(process.Id);
            }

            return null;
        
        }

        #endregion

        #region Public Methods

        public VSProcess(string ROTName, EnvDTE80.DTE2 dte)
        {
            this.ROTName = ROTName;
            this.DTE2 = dte;
        }

        public void Initialize()
        {
            if (m_IsInitializing || IsInitialized)
            {
                return;
            }
            m_IsInitializing = true;

            /**
         * It could be that the DTE2 is currently loading the solution.
         * In that case, accessing the property will result in an error.
         * Give it some time..
         * */
            Solution vsSolution = null;
            int tries = 0;
            while (vsSolution == null && tries <= MaxInitializationRetries)
            {
                tries++;
                try
                {
                    vsSolution = DTE2.Solution;
                    //Test access to solution properties
                    string fullname = vsSolution.FullName;
                }
                catch (Exception e)
                {
                    System.Threading.Thread.Sleep(InitializationRetryMillisecondsTimeout);
                }
            }
            Solution = vsSolution;

            m_IsInitializing = false;
            IsInitialized = true;
        }

        public void ActivateWindow()
        {
            if (!HasDTE)
                return;

            DTE2.MainWindow.Activate();
        }

        public bool OpenSolution(string solutionFilePath)
        {
            if (!HasDTE)
                return false;

            Solution.Open(solutionFilePath);
            IsInitialized = false;
            Initialize();

            return Solution.FullName == solutionFilePath;
        }

        public bool OpenFile(string filePath, int lineNumber = 0, int lineCharacterNumber = -1)
        {
            if (!HasDTE)
                return false;

            EnvDTE.Window w = DTE2.ItemOperations.OpenFile(filePath, EnvDTE.Constants.vsViewKindTextView);
            if (w != null && lineNumber > 0)
            {
                var selection = (EnvDTE.TextSelection)w.Document.Selection;
                if (lineCharacterNumber >= 0)
                {
                    selection.MoveToLineAndOffset(lineNumber, lineCharacterNumber, true);
                    selection.Collapse();
                }
                else
                {
                    selection.GotoLine(lineNumber, true);
                }
                return true;
            }
            return false;
        }

        #endregion
    
        #region Private Fields

        private bool m_IsInitializing;

        #endregion
    }
}