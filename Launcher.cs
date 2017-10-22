using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace VisualStudioLauncher
{
    class Launcher
    {
        public class Parameters
        {
            public string Solution = String.Empty;
            public string File = String.Empty;
            public int Line = 0;
            public int LineCharacter = -1;

            public bool AutofindSolution = true;

            public CommandType Command = CommandType.Default;

            public enum CommandType
            {
                Default,
                Open,
                List,
                LocateSolution
            }
        }

        #region Static Mehods

        public static string LocateParentSolution(string filePath)
        {
            string slnFilePath = null;
            if (File.Exists(filePath))
            {
                string directoryPath = Path.GetDirectoryName(filePath);
                while (directoryPath != null && slnFilePath == null)
                {
                    string[] slnFiles = Directory.GetFiles(directoryPath, "*.sln");
                    if (slnFiles.Length > 0)
                    {
                        slnFilePath = slnFiles[0];
                    }
                    //Go to parent directory
                    directoryPath = Path.GetDirectoryName(directoryPath);
                }
            }
            return slnFilePath;
        }

        #endregion

        #region Public Methods

        public void ParseArguments(string[] arguments)
        {
            Regex argPattern = new Regex(@"^-(\w+)=(.*)$");
            int anonymousIndex = 0;
            foreach (string argument in arguments)
            {
                var argMatch = argPattern.Match(argument);
                if (argMatch.Success)
                {
                    string key = argMatch.Groups[1].Value;
                    string value = argMatch.Groups[2].Value;
                    InjectArgument(key, value);
                }
                else
                {
                    InjectArgument(anonymousIndex, argument);
                    anonymousIndex++;
                }
            }
        }

        public void InjectArgument(string name, string value)
        {
            switch (name)
            {
                case "solution":
                case "s":
                    if (value == "auto")
                        parameters.AutofindSolution = true;
                    else
                        parameters.Solution = Path.GetFullPath(value);
                    break;
                case "file":
                case "f":
                    var match = targetFilePattern.Match(value);
                    if (match.Success)
                    {
                        if(match.Groups[1].Success)
                            parameters.File = match.Groups[1].Value;
                        if (match.Groups[2].Success)
                            parameters.Line = int.Parse(match.Groups[2].Value);
                        if (match.Groups[3].Success)
                            parameters.LineCharacter = int.Parse(match.Groups[3].Value);

                    }
                    else
                    {
                        parameters.File = Path.GetFullPath(value);
                    }
                    break;
                case "line":
                case "l":
                    parameters.Line = int.Parse(value);
                    break;
                case "lineCharacter":
                case "lineChar":
                case "lc":
                    parameters.LineCharacter = int.Parse(value);
                    break;
            }
        }

        public void InjectArgument(int anonymousIndex, string value)
        {
            if (anonymousIndex == 0 && parameters.Command == Parameters.CommandType.Default)
            {
                switch (value)
                {
                    case "open":
                        parameters.Command = Parameters.CommandType.Open;
                        return;
                    case "list":
                    case "ls":
                        parameters.Command = Parameters.CommandType.List;
                        return;
                    case "locate-solution":
                        parameters.Command = Parameters.CommandType.LocateSolution;
                        return;
                }
            }

            if (value.EndsWith(".sln") && string.IsNullOrEmpty(parameters.Solution))
            {
                parameters.Solution = Path.GetFullPath(value);
                return;
            }

            var match = targetFilePattern.Match(value);
            if (match.Success && string.IsNullOrEmpty(parameters.File))
            {
                if (match.Groups[1].Success && string.IsNullOrEmpty(parameters.File))
                    parameters.File = match.Groups[1].Value;
                if (match.Groups[2].Success && parameters.Line == 0)
                    parameters.Line = int.Parse(match.Groups[2].Value);
                if (match.Groups[3].Success && parameters.LineCharacter == -1)
                    parameters.LineCharacter = int.Parse(match.Groups[3].Value);

                return;
            }
        }

        public void ExecuteParameterCommand()
        {
            ExecuteCommand(parameters.Command);
        }

        public void ExecuteCommand(Parameters.CommandType command)
        {
            switch (command)
            {
                case Parameters.CommandType.Default:
                case Parameters.CommandType.Open:
                    UpdateVSProcesses();
                    if (string.IsNullOrEmpty(parameters.Solution) && parameters.AutofindSolution)
                    {
                        parameters.Solution = LocateParentSolution(parameters.File);
                    }
                    var vs = EnsureVSProcessWithSolution(parameters.Solution);
                    if (!string.IsNullOrEmpty(parameters.File))
                    {
                        vs.OpenFile(parameters.File, parameters.Line, parameters.LineCharacter);
                    }
                    vs.ActivateWindow();
                    break;
                case Parameters.CommandType.List:
                    UpdateVSProcesses();
                    Console.WriteLine("Listing running Visual Studio instances:");
                    PrintRunningInstances();
                    break;
                case Parameters.CommandType.LocateSolution:
                    Console.WriteLine($"Searching parent solution for {parameters.File}");
                    string solutionPath = LocateParentSolution(parameters.File);
                    if (String.IsNullOrEmpty(solutionPath))
                    {
                        Console.WriteLine("No solution found");
                    }
                    else
                    {
                        Console.WriteLine("Solution found");
                        Console.WriteLine(solutionPath);
                    }
                    break;
            }
        }

        public void UpdateVSProcesses()
        {
            vsProcesses.Clear();
            VSFinder.GetAllRunningInstances(ref vsProcesses);
        }

        public void PrintRunningInstances()
        {
            VSProcess.Initialize(vsProcesses);
            foreach (var vsProcess in vsProcesses)
            {
                PrintVSInstance(vsProcess);
            }
        }

        public void PrintVSInstance(VSProcess vsProcess)
        {
            if (vsProcess == null || !vsProcess.IsInitialized)
                return;

            Console.WriteLine($"{vsProcess.ROTName}");

            if (vsProcess.HasOpenSolution)
                Console.WriteLine($"\t{vsProcess.Solution.FullName}");
            else
                Console.WriteLine($"\tNo solution open");
        }

        public VSProcess GetVSProcessWithSolution(string solutionFilePath)
        {
            VSProcess.Initialize(vsProcesses);
            foreach (var vsProcess in vsProcesses)
            {
                if (vsProcess.HasOpenSolution && vsProcess.Solution.FullName == solutionFilePath)
                    return vsProcess;
            }
            return null;
        }

        public VSProcess GetVSProcessWithoutSolution()
        {
            VSProcess.Initialize(vsProcesses);
            foreach (var vsProcess in vsProcesses)
            {
                if (!vsProcess.HasOpenSolution)
                    return vsProcess;
            }
            return null;
        }

        public VSProcess GetOrCreateVSProcessForSolution(string solutionFilePath)
        {
            VSProcess vsProcess = null;
            vsProcess = GetVSProcessWithSolution(solutionFilePath);
            if (vsProcess != null)
                return vsProcess;

            vsProcess = GetVSProcessWithoutSolution();
            if (vsProcess != null)
                return vsProcess;
        
            vsProcess = VSProcess.StartNewProcess();
            if (vsProcess != null)
            {
                vsProcess.Initialize();
                vsProcesses.Insert(0, vsProcess);
            }

            return vsProcess;
        }

        public VSProcess EnsureVSProcessWithSolution(string solutionFilePath)
        {
            VSProcess vsProcess = GetOrCreateVSProcessForSolution(solutionFilePath);
            if (vsProcess != null && !vsProcess.HasOpenSolution && !string.IsNullOrEmpty(solutionFilePath))
            {
                vsProcess.OpenSolution(solutionFilePath);
            }
            return vsProcess;
        }

        #endregion

        #region Private Fields

        private Parameters parameters = new Parameters();
        private List<VSProcess> vsProcesses = new List<VSProcess>();

        private Regex targetFilePattern = new Regex(@"^((?:[A-z]:)?[^:*?<>|]+)(?::(\d+))?(?::(\d+))?:*$");

        #endregion
    }
}