using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Task = System.Threading.Tasks.Task;

namespace Smart_Comment
{

    internal sealed class SCInvisible
    {

        public const int CommandId = 256;

        public static readonly Guid CommandSet = new Guid("39e1621e-03d0-496f-886f-79d1d6168ff7");

        private readonly AsyncPackage package;

        DTE dte = Package.GetGlobalService(typeof(SDTE)) as DTE;

        string todayDate = "";
        string projName = "";

        private SCInvisible(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static SCInvisible Instance
        {
            get;
            private set;
        }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SCInvisible(package, commandService);
        }

        public string TodayDatefun()
        {
            DateTime dt = DateTime.Now;
            string strdate = dt.ToString("dd/MM/yyyy");
            strdate = strdate.Replace(".", "/");
            return strdate;
        }

        public string ProjectNamefun()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            projName = Path.GetFileName(dte.Solution.FullName);
            projName = Regex.Replace(projName, ".sln", "");
            return projName;
        }

        string CommentGenerator()
        {
            todayDate = TodayDatefun();
            string projName = ProjectNamefun();
            string output = "// " + projName + ", " + todayDate;
            Clipboard.SetText(output, TextDataFormat.Text);
            return output;
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            System.IServiceProvider serviceProvider = package as System.IServiceProvider;
            IVsStatusbar statusBar = (IVsStatusbar)serviceProvider.GetService(typeof(SVsStatusbar));;
            Assumes.Present(statusBar);

            string result = CommentGenerator();
            statusBar.SetText("Well done! The line is: " + result);

            TextSelection selectedText = (TextSelection)dte.ActiveDocument.Selection;
            
            if(selectedText.Text.Length > 0)
            {
                selectedText.Cut();
                selectedText.Insert(result + " ->");
                selectedText.NewLine();
                selectedText.Insert(Clipboard.GetText());
                selectedText.NewLine();
                selectedText.Insert(result + " <-");
            }
            else 
            {
                selectedText.Insert(Clipboard.GetText());
            }
        }
    }
}
