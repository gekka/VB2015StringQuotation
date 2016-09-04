using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

namespace Gekka.VisualStudio.Extention
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidVSPackage1PkgString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    public sealed class VB2015StringQuotationPackage : Package
    {
        private EnvDTE80.DTE2 dte;
        private EnvDTE80.Events2 event2;
        private EnvDTE80.TextDocumentKeyPressEvents keyPressEvents;

        protected override void Initialize()
        {
            base.Initialize();

            this.dte = (EnvDTE80.DTE2)GetService(typeof(EnvDTE.DTE));
            this.event2 = (EnvDTE80.Events2)this.dte.Events;

            this.keyPressEvents = this.event2.TextDocumentKeyPressEvents;
            this.keyPressEvents.AfterKeyPress += TextDocumentKeyPressEvents_AfterKeyPress;
        }

        void TextDocumentKeyPressEvents_AfterKeyPress(string Keypress, EnvDTE.TextSelection Selection, bool InStatementCompletion)
        {
            try
            {

                if (Keypress == "\r" && Selection.Parent.Language == "Basic")
                {
                    Selection.LineUp();
                    try
                    {
                        Selection.StartOfLine(EnvDTE.vsStartOfLineOptions.vsStartOfLineOptionsFirstText);
                        Selection.EndOfLine(true);

                        string text = Selection.Text;
                        if (text.Length > 0)
                        {
                            int q = 0;
                            foreach (char c in text)
                            {
                                if (c == '\"')
                                {
                                    q++;
                                }
                                else if (c == '\'')
                                {
                                    if ((q & 1) == 0)
                                    {
                                        return;
                                    }

                                }
                            }

                            if ((q & 1) == 1)
                            {
                                Selection.MoveToPoint(Selection.BottomPoint);
                                Selection.Insert("\"");
                            }
                        }

                    }
                    finally
                    {
                        Selection.LineDown();
                        Selection.StartOfLine(EnvDTE.vsStartOfLineOptions.vsStartOfLineOptionsFirstText);

                        //dte.ExecuteCommand("Edit.DocumentFormat");
                    }

                }
            }
            catch
            {
            }
        }
    }
}
