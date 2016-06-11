using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportPreso
{
    public partial class Form1 : Form
    {
        string _fullFilePath = "";
        public Form1()
        {
            InitializeComponent();
            lblMessage.Enabled = false;
        }

        List<SlideInfo> _slideInfos = new List<SlideInfo>();
        private void btnExport_Click(object sender, EventArgs e)
        {

            lblMessage.Enabled = false;
            var result = this.folderBrowser.ShowDialog(this);
            if (result != DialogResult.OK) return;
            var folderPath = FileIO.GetPath(folderBrowser.SelectedPath);

            var ext = new List<string> { ".ppt", ".pptx" };
            var presos = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                 .Where(s => ext.Any(x => s.EndsWith(x))).ToList();

            var pdfext = new List<string> { ".pdf" };
            var pdfs = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                 .Where(s => pdfext.Any(x => s.EndsWith(x))).ToList();

            if (presos.Count() == 0 && pdfs.Count() == 0)
            {
                lblMessage.Text = "no presentations or Pdfs found";
                return;
            }

            var tempDir = Directory.CreateDirectory(folderPath + @"\_temp");

            Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App = null;
            if (presos.Count() != 0)
            {

                try
                {
                    PowerPoint_App = PowerPointEx.ConvertToPPTX(presos, tempDir.FullName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error initializing, close Powerpoint if it is open");
                    return;
                }
            }

            Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;

            string outFile = "";


            foreach (var preso in presos)
            {

                Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(preso, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                var pptName = Path.GetFileName(preso);

                var underscoreIndex = pptName.IndexOf('_');
                var newPptName = pptName.Remove(0, underscoreIndex + 1);

                if (string.IsNullOrEmpty(outFile))
                {
                    outFile = Path.GetFileNameWithoutExtension(newPptName);
                    outFile += ".docx";
                    var outDir = Directory.CreateDirectory(folderPath + @"\_output");
                    _fullFilePath = Path.Combine(outDir.FullName, outFile);
                    WordEx.CreateDoc(_fullFilePath);
                }

                WordEx.AddTitle(newPptName);


                foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in presentation.Slides)
                {

                    bool firstLine = true;

                    //ParseWithOpenXML(slideId,preso);
                    string pptx = Path.GetFileNameWithoutExtension(preso) + ".pptx";
                    string pptxFile = Path.Combine(tempDir.FullName, pptx);
                    _slideInfos.Add(new SlideInfo() { Id = slide.SlideNumber, path = pptxFile });

                    var prevNoteText = "";
                    foreach (var item in slide.Shapes)
                    {
                        //firstLine = true;
                        var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;

                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                        {

                            if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
                            {
                                if (shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Type != MsoBulletType.msoBulletNone)
                                {
                                    List<string> bullets = new List<string>();
                                    foreach (TextRange2 para in shape.TextFrame2.TextRange.Paragraphs)
                                    {
                                        bullets.Add(para.Text);
                                        //var bulletText = "- " + para.Text;
                                        //buffer.AppendLine(bulletText);
                                    }
                                    WordEx.AddBulletList(bullets);
                                    //string text = shape.TextFrame2.TextRange.Text;
                                    //WordEx.AddBulletList(text);

                                }
                                else
                                {
                                    var text = shape.TextFrame2.TextRange.Text;
                                    //slideText += text + " ";
                                    //buffer.AppendLine(text);
                                    if (firstLine)
                                    {
                                        WordEx.AddHeader(text);
                                        firstLine = false;
                                    }
                                    else
                                    {
                                        WordEx.AddText(text);
                                    }
                                }

                            }
                            else if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                                var text = shape.TextFrame.TextRange.Text;
                                //slideText += text + " ";
                                //buffer.AppendLine(text);
                                if (firstLine)
                                {
                                    WordEx.AddHeader(text);
                                    firstLine = false;
                                }
                                else
                                {
                                    WordEx.AddText(text);
                                }
                            }

                            firstLine = false;
                        }

                        if (slide.HasNotesPage == MsoTriState.msoTrue)
                        {
                            //bool processedNotes = false;

                            foreach (var note in slide.NotesPage.Shapes)
                            {
                                //if (processedNotes)
                                //    break;

                                var noteShape = (Microsoft.Office.Interop.PowerPoint.Shape)note;

                                if (noteShape.HasTextFrame == MsoTriState.msoTrue)
                                {

                                    if (noteShape.TextFrame2.HasText == MsoTriState.msoTrue)
                                    {
                                        var text1 = noteShape.TextFrame2.TextRange.Text;
                                        if (text1 == prevNoteText || WordEx.IsNumeric(text1))
                                            continue;

                                        //processedNotes = true; //go to next since notes are duplicated in interop.
                                        if (noteShape.TextFrame2.TextRange.ParagraphFormat.Bullet.Type != MsoBulletType.msoBulletNone)
                                        {

                                            List<string> bullets = new List<string>();
                                            foreach (TextRange2 para in noteShape.TextFrame2.TextRange.Paragraphs)
                                            {
                                                bullets.Add(para.Text);
                                            }
                                            WordEx.AddBulletList(bullets, true);
                                            prevNoteText = noteShape.TextFrame2.TextRange.Text;
                                        }
                                        else
                                        {
                                            var text = noteShape.TextFrame2.TextRange.Text;

                                            WordEx.AddText(text, true);


                                        }

                                    }
                                    else if (noteShape.TextFrame.HasText == MsoTriState.msoTrue)
                                    {
                                        var text = noteShape.TextFrame.TextRange.Text;

                                        WordEx.AddText(text, true);


                                    }
                                }

                            }
                        }
                    }

                }
                presentation.Close();
                //var presoSource = Path.GetFileName(preso);
                //string processedPreso = Path.Combine(processedDir.FullName, presoSource);
                //if (!File.Exists(processedPreso))
                //    File.Copy(preso, processedPreso);

            }
            try
            {
                PowerPointEx.Close(PowerPoint_App);
                AddImages();

                AddPdfs(pdfs);

                WordEx.Save();

                lblMessage.Text = "Created Doc: " + _fullFilePath;
                lblMessage.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void AddPdfs(List<string> pdfs)
        {
            foreach(var pdf in pdfs)
            {
                PdfEx.ConvertToDoc(pdf);
            }
        }

        private void AddImages()
        {
            if (_slideInfos.Count() == 0)
                return;

            var presoFiles = _slideInfos.Select(x => x.path).Distinct();

            foreach (string pptxFile in presoFiles)
            {
                var pptHeader = Path.GetFileNameWithoutExtension(pptxFile);
                var underscoreIndex = pptHeader.IndexOf('_');
                pptHeader = pptHeader.Remove(0, underscoreIndex + 1);

                PresentationDocument ppt = null;

                try
                {
                    ppt = PresentationDocument.Open(pptxFile, true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Powerpoint file needs to be a pptx to extract images.");
                    return;
                }
                PresentationPart presentation = ppt.PresentationPart;


                foreach (var slide in presentation.SlideParts)
                {

                    foreach (ImagePart image in slide.ImageParts)
                    {
                        string header = slide.Slide.InnerText;
                        if (!string.IsNullOrWhiteSpace(header))
                            header = slide.Slide.Descendants<TextBody>().First().InnerText;

                        if (string.IsNullOrWhiteSpace(header))
                        {
                            header = pptHeader;
                        }

                        using (var stream = image.GetStream(FileMode.Open, FileAccess.Read))
                        {
                            var img = Image.FromStream(stream);
                            WordEx.AddImage(img, header, pptHeader);

                        }

                    }

                }

            }

        }


        private void lblMessage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("WINWORD.EXE", "\"" + _fullFilePath + "\"");
        }
    }
}


