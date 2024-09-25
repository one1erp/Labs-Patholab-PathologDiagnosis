
using LSSERVICEPROVIDERLib;
using Patholab_DAL_V1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using ONE1_richTextCtrl;
using forms = System.Windows.Forms;
using Spire.Doc;
using wdUnits = Microsoft.Office.Interop.Word;
using Patholab_Common;


namespace PathologDiagnosis
{

    public partial class PathologDiagnosisWPF : UserControl
    {
        #region Private members

        private INautilusDBConnection _ntlsCon;
        private DataLayer dal;
        private SDG currentSdg = null;
        //   private PathologDiagnosis currentInspection;
        private long currentMicroResultID = -1;
        private string currentMicroResultName = string.Empty;
        private string currentWordFilesPath = string.Empty;
        private List<long> currentResultIDs;
        public bool DEBUG = false;

        Word.Application wordFile = null;
        Word.Document document = null;
        private Object oMissing = System.Reflection.Missing.Value;

        public ONE1_richTextCtrl.RichSpellCtrl richTextMacro;
        public ONE1_richTextCtrl.RichSpellCtrl richTextDiagnosis;
        public RichSpellCtrl richTextMicro;
        public List<ONE1_richTextCtrl.RichSpellCtrl> richSpellCtrls { get; private set; }
        #endregion

        public PathologDiagnosisWPF()
        {
            InitializeComponent();

        }

        #region Public Methods
        public void init(INautilusDBConnection ntlsCon)
        {
            try
            {
                this._ntlsCon = ntlsCon;
                currentResultIDs = new List<long>();

                this.dal = new DataLayer();

                if (DEBUG)
                {
                    dal.MockConnect();
                }
                else
                {
                    dal.Connect(_ntlsCon);
                }

                initRichTextControls();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        public void loadSdg(string sdgName)
        {
            if (!string.IsNullOrEmpty(sdgName))
            {
                try
                {
                    richSpellCtrls.ForEach(x => x.ClearText());

                    currentSdg = dal.FindBy<SDG>(s => s.NAME == sdgName).FirstOrDefault();

                    getResultsID();

                    if (currentResultIDs.Count > 0 && currentSdg.STATUS != "A")
                        OpenRTF();
                    else
                    {
                        richSpellCtrls.ForEach(x => x.Enabled = false);
                        richSpellCtrls.ForEach(x => x.ClearText());
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteLogFile(ex);

                    MessageBox.Show("not a valid SDG." + Environment.NewLine + ex.Message);
                }
            }
        }

        public void saveResults()
        {
            if (currentResultIDs.Count < 1 || currentSdg.STATUS == "A")
                return;

            RESULT resultMacro = dal.FindBy<RESULT>(r => r.NAME.ToLower().Contains("macro") && currentResultIDs.Contains(r.RESULT_ID)).FirstOrDefault();
            RESULT resultDiagnosis = dal.FindBy<RESULT>(r => r.NAME.ToLower().Contains("diagnosis") && currentResultIDs.Contains(r.RESULT_ID)).FirstOrDefault();


            updateResultRTF(richTextMicro, currentMicroResultID);
            updateResultRTF(richTextMacro, resultMacro.RESULT_ID);
            updateResultRTF(richTextDiagnosis, resultDiagnosis.RESULT_ID);
        }

        public void ClearScreen()
        {

            richSpellCtrls.ForEach(x => x.ClearText());

        }
        #endregion


        #region private Methods
        private void getResultsID()
        {
            try
            {
                currentResultIDs.Clear();
                foreach (SAMPLE sample in currentSdg.SAMPLEs)
                {
                    foreach (ALIQUOT aliquot in sample.ALIQUOTs)
                    {
                        foreach (TEST test in aliquot.TESTs)
                        {
                            foreach (RESULT result in test.RESULTs)
                            {
                                currentResultIDs.Add(result.RESULT_ID);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                MessageBox.Show(ex.Message);
            }
        }
        private void OpenRTF()
        {
            RTF_RESULT rtfMicro = null;
            RTF_RESULT rtfMacro = null;
            RTF_RESULT rtfDiagnosis = null;

            try
            {
                dal.RefreshAll();
                RESULT resultMicro = dal.FindBy<RESULT>(r => r.NAME.ToLower().Contains("micro") && currentResultIDs.Contains(r.RESULT_ID)).FirstOrDefault();
                RESULT resultMacro = dal.FindBy<RESULT>(r => r.NAME.ToLower().Contains("macro") && currentResultIDs.Contains(r.RESULT_ID)).FirstOrDefault();
                RESULT resultDiagnosis = dal.FindBy<RESULT>(r => r.NAME.ToLower().Contains("diagnosis") && currentResultIDs.Contains(r.RESULT_ID)).FirstOrDefault();

                try
                {
                    if (resultMicro != null)
                    {
                        rtfMicro = dal.FindBy<RTF_RESULT>(r => r.RTF_RESULT_ID == resultMicro.RESULT_ID).FirstOrDefault();

                        if (rtfMicro != null)
                        {
                            if (string.IsNullOrEmpty(rtfMicro.RTF_TEXT))
                                rtfMicro.RTF_TEXT = string.Empty;
                        }
                        else
                        {
                            // insert new rtf result
                            rtfMicro = new RTF_RESULT();
                            rtfMicro.RTF_RESULT_ID = resultMicro.RESULT_ID;
                            rtfMicro.RTF_TEXT = string.IsNullOrEmpty(resultMicro.FORMATTED_RESULT) ? string.Empty : resultMicro.FORMATTED_RESULT;
                            dal.Add(rtfMicro);

                            dal.SaveChanges();
                            dal.RefreshAll();
                        }

                        currentMicroResultID = resultMicro.RESULT_ID;
                        currentMicroResultName = resultMicro.NAME;
                    }

                    if (resultMacro != null)
                        rtfMacro = dal.FindBy<RTF_RESULT>(r => r.RTF_RESULT_ID == resultMacro.RESULT_ID).FirstOrDefault();
                    if (resultDiagnosis != null)
                        rtfDiagnosis = dal.FindBy<RTF_RESULT>(r => r.RTF_RESULT_ID == resultDiagnosis.RESULT_ID).FirstOrDefault();
                }
                catch (Exception)
                {

                    MessageBox.Show("result not found.");
                }

                if (rtfMacro != null)
                {
                    initControlText(richTextMacro, rtfMacro);
                }

                if (rtfDiagnosis != null)
                {
                    initControlText(richTextDiagnosis, rtfDiagnosis);
                }

                if (rtfMicro != null)
                {
                    currentWordFilesPath = getWordFilesPath();

                    initControlText(richTextMicro, rtfMicro);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                MessageBox.Show(ex.Message);
            }
        }
        private string getWordFilesPath()
        {
            PHRASE_HEADER pHeader = dal.FindBy<PHRASE_HEADER>(header => header.NAME.ToLower().Equals("system parameters")).FirstOrDefault();
            if (pHeader == null) return null;

            PHRASE_ENTRY pEntry = dal.FindBy<PHRASE_ENTRY>(entry => entry.PHRASE_NAME.ToLower().Equals("worklist word files")).FirstOrDefault();
            if (pEntry == null) return null;

            return Path.Combine(pEntry.PHRASE_DESCRIPTION, currentMicroResultID + "_" + currentMicroResultName.Replace(" ", "_") + ".rtf");
        }
        private void initControlText(ONE1_richTextCtrl.RichSpellCtrl richTextCtrl, RTF_RESULT result)
        {
            if (!string.IsNullOrEmpty(result.RTF_TEXT))
                richTextCtrl.SetRtf(result.RTF_TEXT.Trim());
        }
        private void initRichTextControls()
        {
            Logger.WriteLogFile("before initRichTextControls");
            try
            {
                if (richTextMicro == null) richTextMicro = new ONE1_richTextCtrl.RichSpellCtrl();
                if (richTextMacro == null) richTextMacro = new ONE1_richTextCtrl.RichSpellCtrl();
                if (richTextDiagnosis == null) richTextDiagnosis = new ONE1_richTextCtrl.RichSpellCtrl();

                richSpellCtrls = new List<ONE1_richTextCtrl.RichSpellCtrl>();
                richSpellCtrls.Add(richTextMacro);
                richSpellCtrls.Add(richTextDiagnosis);
                richSpellCtrls.Add(richTextMicro);


                foreach (var item in richSpellCtrls)
                {
                    item.SetDefaultSpelling();
                    item.RightToLeft = forms.RightToLeft.Yes;

                }

                winformsHostMacro.Child = richTextMacro;
                winformsHostDiagnos.Child = richTextDiagnosis;
                winformsHostMicro.Child = richTextMicro;

                richTextMicro.DocumentBody.MouseDoubleClick += this.rtbDocument_MouseDoubleClick;
            }
            catch (Exception ex) {
                Logger.WriteLogFile($"while initRichTextControls: {ex}");
            }
        }
        // button to open micro as word process
        private void rtbDocument_MouseDoubleClick(object sender, forms.MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Wait;

            saveRtfTextAsWord(richTextMicro.GetRtf(), currentWordFilesPath);
            deleteFirstRow(currentWordFilesPath);
            openFile(currentWordFilesPath);

            Mouse.OverrideCursor = null;
        }
        private void openFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    Mouse.OverrideCursor = null;

                    Process wordProcess = Process.Start(filePath);
                    wordProcess.WaitForExit();

                    Mouse.OverrideCursor = Cursors.Wait;

                    wordProcess.Close();
                    wordProcess.Dispose();
                    ProcessExited();

                    Mouse.OverrideCursor = null;
                }
                else
                {
                    MessageBox.Show("Invalid path. Try setting a valid path in 'system parameters' phrase.");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                MessageBox.Show(ex.Message);
            }
        }
        private void ProcessExited()
        {
            try
            {
                updateResultRTF(richTextMicro, currentMicroResultID, true);
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                MessageBox.Show(ex.Message);
            }
        }
        private void updateResultRTF(ONE1_richTextCtrl.RichSpellCtrl richText, long resultID, bool hasWordFile = false)
        {
            try
            {
                RTF_RESULT rtf = dal.FindBy<RTF_RESULT>(r => r.RTF_RESULT_ID == resultID).FirstOrDefault();
                string rtfString;
                string text;

                if (File.Exists(currentWordFilesPath) && hasWordFile)
                    rtfString = File.ReadAllText(currentWordFilesPath);
                else
                    rtfString = richText.GetRtf();

                text = rtfStringToText(rtfString);
                text = text.Substring(0, text.Length > 4000 ? 4000 : text.Length);

                if (rtf != null)
                {
                    rtf.RTF_TEXT = rtfString;
                }
                else
                {
                    var newRTF = new RTF_RESULT();
                    newRTF.RTF_RESULT_ID = resultID;
                    newRTF.RTF_TEXT = rtfString;
                    dal.Add(newRTF);
                }

                RESULT result = dal.FindBy<RESULT>(r => r.RESULT_ID == resultID).FirstOrDefault();
                result.FORMATTED_RESULT = text;

                dal.SaveChanges();
                dal.RefreshAll();

                initControlText(richText, rtf);
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);

                if (ex.Message.Contains("ORA-00942"))
                {

                    MessageBox.Show("Logged in user can't edit database");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private string rtfStringToText(string RtfString)
        {
            ONE1_richTextCtrl.RichSpellCtrl rtBox = new ONE1_richTextCtrl.RichSpellCtrl();

            rtBox.SetRtf(RtfString);

            return rtBox.GetOriginalText();
        }
        private void deleteFirstRow(string filePath)
        {
            try
            {
                wordFile = new Word.Application();
                document = wordFile.Documents.Open(filePath);
                var range = document.Content;
                if (range.Find.Execute("Evaluation Warning: The document was created with Spire.Doc for .NET."))
                {
                    range.Expand(wdUnits.WdUnits.wdSentence); // or change to .wdLine or .wdSentence or .wdParagraph
                    range.Delete();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (document != null)
                    document.Close(ref oMissing, ref oMissing, ref oMissing);
                if (wordFile != null)
                    wordFile.Quit(ref oMissing, ref oMissing, ref oMissing);
            }
        }
        #endregion









        private static void saveRtfTextAsWord(string rtf, string filePathAndName)
        {
            Document doc = null;

            try
            {
                doc = new Document();

                TextReader tr = new StringReader(rtf);

                doc.LoadRtf(tr);
                doc.SaveToFile(filePathAndName, FileFormat.Rtf);

            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);

                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    doc.Dispose();
                }
            }
        }





        private void winformsHostDiagnos_ChildChanged(object sender, forms.Integration.ChildChangedEventArgs e)
        {

        }
    }


}
