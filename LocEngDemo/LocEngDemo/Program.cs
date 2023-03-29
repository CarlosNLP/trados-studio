using Sdl.Core.Globalization;
using Sdl.ProjectAutomation.Core;
using Sdl.ProjectAutomation.FileBased;
using System.Globalization;
using System;

namespace RWSApiSample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Creating the variables for the project
            var projectName = "LocEngDemo";
            var projectFolder = @"C:/Docs/Localization Academy/Trados Studio/Project/";
            var template = @"C:/Docs/Localization Academy/Trados Studio/Template/LocEngTemplate.sdltpl";
            var sourcePath = @"C:/Docs/Localization Academy/Trados Studio/Source/";

            // Defining the project information
            ProjectInfo projectInfo = new ProjectInfo()
            {
                Name = projectName,
                LocalProjectFolder = projectFolder,
            };

            // Creating the Studio project with template and project information
            Console.WriteLine("Creating the Trados Studio project...");
            ProjectTemplateReference templatePath = new ProjectTemplateReference(template);
            FileBasedProject studioProject = new FileBasedProject(projectInfo, templatePath);

            // Adding the source files from folder in recursive mode
            Console.WriteLine("Adding the source files to the project...");
            studioProject.AddFolderWithFiles(sourcePath, true);

            // Preparing the source files with the necessary batch tasks
            Console.WriteLine("Preparing the source files for translation...");
            ProjectFile[] sourceFiles = studioProject.GetSourceLanguageFiles();
            AutomaticTask scanTask = studioProject.RunAutomaticTask(sourceFiles.GetIds(), AutomaticTaskTemplateIds.Scan);
            AutomaticTask convertTask = studioProject.RunAutomaticTask(sourceFiles.GetIds(), AutomaticTaskTemplateIds.ConvertToTranslatableFormat);
            AutomaticTask copyTask = studioProject.RunAutomaticTask(sourceFiles.GetIds(), AutomaticTaskTemplateIds.CopyToTargetLanguages);
            studioProject.Save();

            // Pretranslating the files
            ProjectFile[] targetFiles = studioProject.GetTargetLanguageFiles();
            Console.WriteLine("Running pretranslation...");
            AutomaticTask pretranslateTask = studioProject.RunAutomaticTask(targetFiles.GetIds(), AutomaticTaskTemplateIds.PreTranslateFiles);

            // Analyzing the files
            Console.WriteLine("Running analysis and saving as Excel format...");
            var targetLanguages = new[]
            {
                new Language(new CultureInfo("es-ES")),
                new Language(new CultureInfo("fr-FR")),
                new Language(new CultureInfo("it-IT")),
                new Language(new CultureInfo("el-GR"))
            };

            for (var i = 0; i < targetLanguages.Length; i++)
            {
                try
                {
                    ProjectFile[] analysisTargetFiles = studioProject.GetTargetLanguageFiles(targetLanguages[i]);
                    AutomaticTask analyzeTask = studioProject.RunAutomaticTask(analysisTargetFiles.GetIds(), AutomaticTaskTemplateIds.AnalyzeFiles);
                    var reportId = analyzeTask.Reports[0].Id;
                    studioProject.SaveTaskReportAs(reportId, projectFolder + @"Reports/Report_" + targetLanguages[i].ToString() + ".xlsx", ReportFormat.Excel);
                }

                catch (Exception)
                {
                    Console.WriteLine("The analysis task for {0} threw an error. The language is skipped.", targetLanguages[i].ToString());
                }
            }

            // Generating the translation packages
            Console.WriteLine("Generating the Trados Studio packages for translation...");

            for (var i = 0; i < targetLanguages.Length; i++)
            {
                try
                {
                    ProjectFile[] packageFiles = studioProject.GetTargetLanguageFiles(targetLanguages[i]);
                    ManualTask packageTask = studioProject.CreateManualTask("Translate", "Translator", DateTime.Now.AddDays(3), packageFiles.GetIds());
                    ProjectPackageCreation projectPackage = studioProject.CreateProjectPackage(packageTask.Id, "Sample Package", "Comment", GetPackageOptions());
                    studioProject.SavePackageAs(projectPackage.PackageId, projectFolder + @"Packages\" + projectName + "_" + targetLanguages[i].ToString() + ".sdlppx");
                }

                catch (Exception)
                {
                    Console.WriteLine("The package generation for {0} threw an error. The language is skipped.", targetLanguages[i].ToString());
                }
            }

            // Printing messages to let the user know the process is now completed
            Console.WriteLine();
            Console.WriteLine("The process is now completed. The project is saved in the below path:");
            Console.WriteLine(projectFolder);

        }

        public static ProjectPackageCreationOptions GetPackageOptions()
        {
            ProjectPackageCreationOptions options = new ProjectPackageCreationOptions();
            options.IncludeAutoSuggestDictionaries = false;
            options.IncludeMainTranslationMemories = true;
            options.RemoveServerBasedTranslationMemories = false;
            options.IncludeTermbases = true;
            options.RemoveAutomatedTranslationProviders = false;
            options.RecomputeAnalysisStatistics = false;
            options.ProjectTranslationMemoryOptions = ProjectTranslationMemoryPackageOptions.CreateNew;

            return options;
        }
    }
}
