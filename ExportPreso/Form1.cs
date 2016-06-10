using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
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
        public struct SlideInfo
        {
            public int Id;
            public string path;
        }
        List<SlideInfo> _slideInfos = new List<SlideInfo>();
        private void btnExport_Click(object sender, EventArgs e)
        {


            lblMessage.Enabled = false;
            var result = this.folderBrowser.ShowDialog(this);
            if (result != DialogResult.OK) return;
            var folderPath = FileIO.GetPath(folderBrowser.SelectedPath);

            //string folderPath = @"C:\Projects\ExportPreso\Test Presos\FolderTest";
            //string filePath = folderPath + @"\1_PY620C class 6 -t tests.ppt";

            var ext = new List<string> { ".ppt", ".pptx" };
            var presos = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                 .Where(s => ext.Any(x => s.EndsWith(x)));

            if (presos.Count() == 0)
            {
                lblMessage.Text = "no presentations found";
                return;
            }

            var processedDir = Directory.CreateDirectory(folderPath + @"\_processed");
            //PowerPointEx pp = new PowerPointEx();
            //bool ppReady = pp.EnsurePowerPoint();

            Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App;
            try
            {
                PowerPoint_App = new Microsoft.Office.Interop.PowerPoint.Application();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initializing, close Powerpoint if it is open");
                return;
            }
            Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;

            //StringBuilder buffer = new StringBuilder();
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
                    int slideId = slide.SlideID;
                    //buffer.AppendLine();
                    //WordEx.AddBlankLine();
                    bool firstLine = true;

                    //ParseWithOpenXML(slideId,preso);
                    _slideInfos.Add(new SlideInfo() { Id = slideId, path = preso });
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
                                        if (text1 == prevNoteText||WordEx.IsNumeric(text1))
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
                var presoSource = Path.GetFileName(preso);
                string processedPreso = Path.Combine(processedDir.FullName, presoSource);
                if (!File.Exists(processedPreso))
                    File.Copy(preso, processedPreso);

            }
            try
            {
                PowerPoint_App.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PowerPoint_App);
                Process[] pros = Process.GetProcesses();
                for (int i = 0; i < pros.Count(); i++)
                {
                    if (pros[i].ProcessName.ToLower().Contains("powerpnt"))
                    {
                        pros[i].Kill();
                    }
                }
                //ParseWithOpenXML(); //maybe later to get images

                WordEx.Save();

                lblMessage.Text = "Created Doc: " + _fullFilePath;
                lblMessage.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ParseWithOpenXML()
        {
            var preso = _slideInfos[0].path;
            var id = _slideInfos[0].Id;

            PresentationDocument ppt = null;

            try
            {
                ppt = PresentationDocument.Open(preso, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Powerpoint file needs to be a pptx to extract images.");
                return;
            }
            PresentationPart presentation = ppt.PresentationPart;
            // get the SlideIdList
            var items = presentation.Presentation.SlideIdList;

            // enumerate over that list
            foreach (SlideId item in items)
            {
                // get the "Part" by its "RelationshipId"
                var part = presentation.GetPartById(item.RelationshipId);

                // this part is really a "SlidePart" and from there, we can get at the actual "Slide"
                var slide = (part as SlidePart).Slide;
                string xmlSlideId = presentation.GetIdOfPart(slide.SlidePart);
                // do more stuff with your slides here!
                //slide.rel
            }

            foreach (var slide in presentation.SlideParts)
            {

                foreach (var part in slide.Parts)
                {
                    var pt = part.RelationshipId;
                    //if(part.part)
                }
                foreach (var part in slide.SlideParts)
                {
                    var pt = part.SlideCommentsPart;
                    //if(part.part)
                }
                foreach (ImagePart image in slide.ImageParts)
                {
                    // var Image = ImagePartType;
                    var stream = image.GetStream();
                }


            }
            //var slide = presentation.GetPartsOfType<SlidePart>().FirstOrDefault();

            //var imagePart = slide.GetPartsOfType<ImagePart>().FirstOrDefault();

            //var stream = imagePart.GetStream();

            //var img = Image.FromStream(stream);
        }
        ///////////




        private void lblMessage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("WINWORD.EXE", "\"" + _fullFilePath + "\"");
        }
    }
}


