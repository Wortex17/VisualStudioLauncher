using System.Collections.Generic;

namespace VisualStudioLauncher
{
    class VSFinder
    {
        public static void GetAllRunningInstances(ref List<VSProcess> runningVSInstances)
        {
            IDictionary<string, object> runningObjects = ProcessFinder.GetRunningObjectTables();

            foreach (string objectName in runningObjects.Keys)
            {
                if (objectName.StartsWith("!VisualStudio.DTE"))
                {
                    var vsProcess = VSProcess.FromROT(objectName, runningObjects[objectName]);
                    if (vsProcess != null)
                    {
                        runningVSInstances.Add(vsProcess);
                    }
                }
            }
        }

        public static VSProcess GetInstanceByProcessId(int processId)
        {
            IDictionary<string, object> runningObjects = ProcessFinder.GetRunningObjectTables();

            foreach (string objectName in runningObjects.Keys)
            {
                if (objectName.StartsWith("!VisualStudio.DTE") && objectName.EndsWith($"{processId}"))
                {
                    var vsProcess = VSProcess.FromROT(objectName, runningObjects[objectName]);
                    if (vsProcess != null)
                    {
                        return vsProcess;
                    }
                }
            }

            return null;
        }
    
    }
}