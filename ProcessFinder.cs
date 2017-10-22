using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace VisualStudioLauncher
{
    class ProcessFinder
    {
        public static IDictionary<string, object> GetRunningObjectTables()
        {
            IDictionary<string, object> rotTable = new Dictionary<string, object>();

            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            IMoniker[] monikers = new IMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            IntPtr numberFetched = IntPtr.Zero;

            while (monikerEnumerator.Next(1, monikers, numberFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);
                Marshal.ReleaseComObject(ctx);

                object runningObjectValue;
                runningObjectTable.GetObject(monikers[0], out runningObjectValue);

                if (!rotTable.ContainsKey(runningObjectName))
                    rotTable.Add(runningObjectName, runningObjectValue);
            }

            return rotTable;
        }

        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);
        [DllImport("ole32.dll")]
        public static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
    }
}