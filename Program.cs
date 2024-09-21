/////////////////////////////////////////////////////////////////////////////////////////
// Titel: Template filler with commandline interface and Miniword / Miniexcel backend
// Author: Patrick Beck
// date: 2024
/////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Text.Json;
using CommandLine;
using DocumentFormat.OpenXml.Wordprocessing;
using MiniSoftware;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System.Data;
using System.Xml.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Bibliography;
using MiniExcelLibs;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;

namespace DocumentCreator
{
    class Program
    {

        enum ExitCode : int
        {
            Success = 0,
            InvalidParameter= 328,
            InvalidFilename = 2,
            FileAccessDenied = 5,
            UnknownError = 10
        }

        // Commands of one set can run at a time. So only "InputFile" or "Commandline" will be parsed successfully. Commands without setname can be run in both sets.
        public class Options
        {
            [Option('i', "input", SetName = "InputFile", Required = true, Separator = ',', HelpText = "Set input file(s) for variables if no commandline parameters are specified. More than one input file are separated by space.")]
            public IEnumerable<string> InputFile { get; set; }

            [Option('m', "inputmode", SetName = "InputFile", Required = true, HelpText = "Set input file(s) insert mode. Can be placeholder (p), list (l) or table (t). Has to be selected for every input file. For example 'p p p' for three files or 't' for one file ")]
            public IEnumerable<string> InputMode { get; set; }

            [Option('n', "inputname", SetName = "InputFile", Required = true, HelpText = "For every file a variable name can be selected. The data access pattern is {{InputName.ColumnName}} in the Word file")]
            public IEnumerable<string> InputName { get; set; }

            [Option('t', "template", Required = true, HelpText = "Set template file for file creation")]
            public string TemplateFile { get; set; }

            [Option('o', "output", Required = true, HelpText = "Set output file")]
            public string OutputFile { get; set; }

            [Option('p', "placeholder", SetName = "Commandline", Required = true, Separator = ',', HelpText = "If no input file is specified placeholder and variables can be passed by command line interface. Placeholder and variables has to be the same size. Separated by space. Insert mode placeholder.")]
            public IEnumerable<string> Placeholder { get; set; }

            [Option('v', "variables", SetName = "Commandline", Required = true, Separator = ',', HelpText = "If no input file is specified placeholder and variables can be passed by command line interface. Placeholder and variables has to be the same size. Separated by space. Inset mode placeholder.")]
            public IEnumerable<string> Variables { get; set; }   
        }

        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed(Run);
        }

        private static void Run(Options cmdParameter)
        {
            // Store the replacment information
            var tagDict = new Dictionary<string, object>();

            var ListInputFiles = cmdParameter.InputFile.ToList();
            var ListInputNames = cmdParameter.InputName.ToList();
            var ListInputModes = cmdParameter.InputMode.ToList();

            // SetName InputFile selected
            if (ListInputFiles.Count > 0)
            {
                // Every Input file need an correct formed input mode
                if (ListInputFiles.Count != ListInputNames.Count || ListInputFiles.Count != ListInputModes.Count)
                {
                    Console.WriteLine("Error: InputFiles, InputNames and InputModes has to match");
                    // Exit with invalid parameter
                    System.Environment.Exit((int)ExitCode.InvalidParameter);
                }

                // combine all elements in one list
                IEnumerable<(string file, string name, string mode)> combineFileNameList = cmdParameter.InputFile.Zip(cmdParameter.InputName, cmdParameter.InputMode);

                foreach (var element in combineFileNameList)
                {
                    try
                    {
                        // open input file
                        using (var stream = File.Open(element.file, FileMode.Open, FileAccess.Read))
                        {
                            IExcelDataReader reader;

                            // This is required to parse strings in binary BIFF2-5 Excel documents encoded with DOS-era code pages.
                            // These encodings are registered by default in the full .NET Framework, but not on .NET Core and .NET 5.0 or later.
                            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                            reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                            var conf = new ExcelDataSetConfiguration
                            {
                                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                                {
                                    UseHeaderRow = true
                                }
                            };

                            var dataSet = reader.AsDataSet(conf);

                            // Now you can get data from each sheet by its index or its "name"
                            // first sheet
                            var dataTable = dataSet.Tables[0];

                            switch (element.mode)

                            {
                                case "t":
                                    //datastorage Method for table insert in MiniWord is Dict filled with list of dicts
                                    var ListOfDicts = new List<Dictionary<string, object>>();

                                    foreach (DataRow dr in dataTable.Rows)
                                    {
                                        // for storing a row - the column key is always the same
                                        var tempDict = new Dictionary<string, object>();
                                        for (int i = 0; i < dataTable.Columns.Count; i++)
                                        {
                                            tempDict.Add(dataTable.Columns[i].ColumnName, dr[dataTable.Columns[i].ColumnName].ToString());
                                        }
                                        ListOfDicts.Add(tempDict);
                                    }
                                    tagDict.Add(element.name, ListOfDicts);
                                    break;

                                case "l":
                                    //datastorage Method for list insert in MiniWord is Dict with list elements

                                    foreach (DataColumn column in dataTable.Columns)
                                    {
                                        var ListOfElements = new List<string>();
                                        foreach (DataRow row in dataTable.Rows)
                                        {
                                            ListOfElements.Add(row[column.ColumnName].ToString());
                                        }
                                        // the tag syntax in the word file should be identical to the table so we add the Inputname.
                                        // So its also possible to have the same column names in different files
                                        tagDict.Add(element.name + '.' + column.ColumnName, ListOfElements);
                                    }
                                    break;

                                case "p":
                                    //datastorage Method for tag insert is only a dict :)

                                    foreach (DataRow dr in dataTable.Rows)
                                    {
                                        tagDict.Add(element.name + '.' + dr[0].ToString(), dr[1].ToString());
                                    }
                                    break;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error: File not found: {0}", element.file);
                        // Exit with FileNotFound
                        System.Environment.Exit((int)ExitCode.InvalidFilename);
                    }
                }
            }


            // convert only for checking with count - not neccesary for adding to dict
            var ListPlaceholder = cmdParameter.Placeholder.ToList();
            var ListVariables = cmdParameter.Variables.ToList();

            // SetName Commandline selected
            if (ListVariables.Count > 0)
            {

                // check if lists has the same size
                if (ListPlaceholder.Count != ListVariables.Count)
                {
                    Console.WriteLine("Error: Wrong length of placeholder and variable.");
                    // Exit with invalid parameter
                    System.Environment.Exit((int)ExitCode.InvalidParameter);
                }

                // combine two list to one for easy add to dict
                var combindList = cmdParameter.Placeholder.Zip(cmdParameter.Variables, (p, v) => new { placeholder = p, variable = v });

                foreach (var element in combindList)
                    tagDict.Add(element.placeholder, element.variable);
            }

            try
            {
                // load template file (given by commandline parameters)
                var template = File.ReadAllBytes(cmdParameter.TemplateFile);

                try
                {
                    // create output file

                    
                    // Dirty check for file extension - can be faulty ...
                    if (Path.GetExtension(cmdParameter.OutputFile) == ".docx")
                    {
                        MiniWord.SaveAsByTemplate(cmdParameter.OutputFile, template, tagDict);
                    }
                    else
                    {
                        MiniExcel.SaveAsByTemplate(cmdParameter.OutputFile, template, tagDict);
                    }


                    Console.WriteLine("Output file: {0} sucessfully created", cmdParameter.OutputFile);
                    // Exit with no error
                    System.Environment.Exit((int)ExitCode.Success);
                }

                catch (FileNotFoundException e)
                {
                    Console.WriteLine("Error: File not found: {0}", cmdParameter.OutputFile);
                    // Exit with FileNotFound
                    System.Environment.Exit((int)ExitCode.InvalidFilename);
                }
                catch (IOException e)
                {
                    Console.WriteLine("Error: File access locked: {0}", cmdParameter.OutputFile);
                    // Exit with File access denied
                    System.Environment.Exit((int)ExitCode.FileAccessDenied);
                }

                catch (Exception e)
                {
                    Console.WriteLine("Error: Something went wrong :( - {0}", e);
                    // Exit with File access denied
                    System.Environment.Exit((int)ExitCode.UnknownError);
                }
            }

            catch (FileNotFoundException e)
            {
                Console.WriteLine("Error: File not found: {0}", cmdParameter.TemplateFile);
                // Exit with FileNotFound
                System.Environment.Exit((int)ExitCode.InvalidFilename);
            }
            catch (IOException e)
            {
                Console.WriteLine("Error: File access locked: {0}", cmdParameter.TemplateFile);
                // Exit with File access denied
                System.Environment.Exit((int)ExitCode.FileAccessDenied);
            }
        }
    }
}