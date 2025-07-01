﻿using alxnbl.OneNoteMdExporter.Helpers;
using alxnbl.OneNoteMdExporter.Infrastructure;
using alxnbl.OneNoteMdExporter.Models;
using Microsoft.Office.Interop.OneNote;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;

namespace alxnbl.OneNoteMdExporter.Services.Export
{
    /// <summary>
    /// Base class for Export Service. 
    /// Contains all shared logic between exporter of different formats.
    /// Abstract methods needs to be implemented by each exporter
    /// </summary>
    public abstract class ExportServiceBase : IExportService
    {
        protected abstract string ExportFormatCode { get; }

        protected static string GetNotebookFolderPath(Notebook notebook)
            => Path.Combine(notebook.ExportFolder, notebook.GetNotebookPath());

        /// <summary>
        /// Return location in the export folder of an attachment file
        /// </summary>
        /// <param name="page"></param>
        /// <param name="attachId">Id of the attachment</param>
        /// <param name="oneNoteFilePath">Original file path of the file in OneNote</param>
        /// <returns></returns>
        protected abstract string GetAttachmentFilePath(Attachement attachment);

        /// <summary>
        /// Get the md reference to the attachment
        /// </summary>
        /// <param name="attachment"></param>
        /// <returns></returns>
        protected abstract string GetAttachmentMdReference(Attachement attachment);

        protected abstract string GetResourceFolderPath(Page node);

        protected abstract string GetPageMdFilePath(Page page);


        public NotebookExportResult ExportNotebook(Notebook notebook, string sectionNameFilter = "", string pageNameFilter = "")
        {
            notebook.ExportFolder = @$"{Localizer.GetString("ExportFolder")}\{ExportFormatCode}\{notebook.GetNotebookPath()}-{DateTime.Now:yyyyMMdd HH-mm}";
            CleanUpFolder(notebook);

            // Initialize hierarchy of the notebook from OneNote APIs
            try
            {
                OneNoteApp.Instance.FillNodebookTree(notebook);
            }
            catch (Exception ex)
            {
                return new NotebookExportResult
                {
                    NoteBookExportErrorCode = "ErrorDuringNotebookProcessingNbTree",
                    NoteBookExportErrorMessage = string.Format(Localizer.GetString("ErrorDuringNotebookProcessingNbTree"),
                        notebook.Title, notebook.Id, ex.Message)
                };
            }

            return ExportNotebookInTargetFormat(notebook, sectionNameFilter, pageNameFilter);
        }

        public abstract NotebookExportResult ExportNotebookInTargetFormat(Notebook notebook, string sectionNameFilter = "", string pageNameFilter = "");

        private static void CleanUpFolder(Notebook notebook)
        {
            // Cleanup Notebook export folder
            DirectoryHelper.ClearFolder(GetNotebookFolderPath(notebook));

            // Cleanup temp folder
            DirectoryHelper.ClearFolder(GetTmpFolder(notebook));
        }

        protected abstract void PrepareFolders(Page page);

        protected static string GetTmpFolder(Node node)
            => Path.Combine(Path.GetTempPath(), node.GetNotebookPath());

        /// <summary>
        /// Export a Page and its attachments
        /// </summary>
        /// <param name="page"></param>
        /// <param name="retry">True if the execution is caused by a retry after an error on the page</param>
        /// <returns>True if the export finished with success</returns>
        protected bool ExportPage(Page page, bool retry = false)
        {
            try
            {
                OneNoteApp.Instance.GetPageContent(page.OneNoteId, out var xmlPageContentStr, PageInfo.piBinaryDataFileType);
                
                
                //// Generate "onenote:" link to extract section-id and page-id (eg: "Pag 2")
                //OneNoteApp.Instance.GetHyperlinkToObject("{095C0BE4-7C7A-4CA5-92D3-9835FDD105E5}{1}{E19538458627694935540520157643112716112073471}", null, out var test);


                // Alternative : return page content without binaries
                //oneNoteApp.GetHierarchy(page.OneNoteId, HierarchyScope.hsChildren, out var xmlAttach);

                var xmlPageContent = XDocument.Parse(xmlPageContentStr).Root;
                var ns = xmlPageContent.Name.Namespace;
                page.Author = xmlPageContent.Element(ns + "Title")?.Element(ns + "OE")?.Attribute("author")?.Value ?? "unknown";
                ProcessPageAttachments(ns, page, xmlPageContent);

                // Suffix page title
                EnsurePageUniquenessPerSection(page);

                // Make various OneNote XML fixes before page export
                page.OverrideOneNoteId = PageXmlPreProcessing(xmlPageContent);

                // Register page and section mappings for link conversion
                var pagePath = page.GetPageFileRelativePath(AppSettings.MdMaxFileLength);
                
                // Generate programmatic ID for the page
                string pageProgrammaticId = null;
                try
                {
                    OneNoteApp.Instance.GetHyperlinkToObject(page.OneNoteId, null, out string pageLink);
                    var pageIdMatch = Regex.Match(pageLink, @"page-id=\{([^}]+)\}", RegexOptions.IgnoreCase);
                    if (pageIdMatch.Success)
                    {
                        pageProgrammaticId = pageIdMatch.Groups[1].Value;
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning($"Failed to generate programmatic ID for page {page.Title}: {ex.Message}");
                }
                
                ConverterService.RegisterPageMapping(page.OneNoteId, pageProgrammaticId, pagePath, page.Title);

                // Generate programmatic ID for the section
                string sectionProgrammaticId = null;
                try
                {
                    if (page.Parent?.OneNoteId != null)
                    {
                        OneNoteApp.Instance.GetHyperlinkToObject(page.Parent.OneNoteId, null, out string sectionLink);
                        var sectionIdMatch = Regex.Match(sectionLink, @"section-id=\{([^}]+)\}", RegexOptions.IgnoreCase);
                        if (sectionIdMatch.Success)
                        {
                            sectionProgrammaticId = sectionIdMatch.Groups[1].Value;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning($"Failed to generate programmatic ID for section {page.Parent?.Title}: {ex.Message}");
                }
                
                if (page.Parent?.OneNoteId != null)
                {
                    ConverterService.RegisterSectionMapping(page.Parent.OneNoteId, sectionProgrammaticId, page.Parent.GetPath(AppSettings.MdMaxFileLength), page.Parent.Title);
                }

                var docxFileTmpFile = Path.Combine(GetTmpFolder(page), page.Id + ".docx");

                if (File.Exists(docxFileTmpFile))
                    File.Delete(docxFileTmpFile);

                PrepareFolders(page);

                Log.Debug($"{page.OneNoteId}: start OneNote docx publish");
                if (page.OverrideOneNoteId != null)
                    Log.Debug($"Actually using temporary page ${page.OverrideOneNoteId}");

                // Request OneNote to export the page into a DocX file
                OneNoteApp.Instance.Publish(page.OverrideOneNoteId ?? page.OneNoteId, Path.GetFullPath(docxFileTmpFile), PublishFormat.pfWord);

                Log.Debug($"{page.OneNoteId}: success");

                if (AppSettings.Debug || AppSettings.KeepOneNoteTempFiles)
                {
                    // If debug mode enabled, copy the page docx file next to the page md file
                    var docxFilePath = Path.ChangeExtension(GetPageMdFilePath(page), "docx");
                    File.Copy(docxFileTmpFile, docxFilePath);
                }

                // Convert docx file into Md using PanDoc
                var pageMd = ConverterService.ConvertDocxToMd(page, docxFileTmpFile, GetTmpFolder(page));

                if (AppSettings.Debug)
                {
                    // And write Pandoc markdown file
                    var mdPanDocFilePath = Path.ChangeExtension(GetPageMdFilePath(page), "pandoc.md");
                    File.WriteAllText(mdPanDocFilePath, pageMd);
                }

                File.Delete(docxFileTmpFile);

                // Copy images extracted from DocX to Export folder and add them in the list of attachments of the page
                try
                {
                    ExtractImagesToResourceFolder(page, ref pageMd);
                }
                catch (COMException ex)
                {
                    if (ex.Message.Contains("0x800706BE"))
                    {
                        LogError(page, ex, Localizer.GetString("ErrorWhileStartingOnenote"));
                    }
                    else
                        LogError(page, ex, string.Format(Localizer.GetString("ErrorDuringOneNoteExport"), ex.Message));
                }
                catch (Exception ex)
                {
                    LogError(page, ex, Localizer.GetString("ErrorImageExtract"));
                }

                // Export all file attachments and get updated page markdown including md reference to attachments
                ExportPageAttachments(page, ref pageMd);

                // Apply post processing to Page Md content
                ConverterService.PageMdPostConversion(ref pageMd);

                // Apply post processing specific to an export format
                pageMd = FinalizePageMdPostProcessing(page, pageMd);

                WritePageMdFile(page, pageMd);

                return true;
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("0x800706BE"))
                {
                    LogError(page, ex, string.Format(Localizer.GetString("ErrorDuringPageProcessingIsOneNoteRunning"), page.TitleWithPageLevelTabulation, page.Id, ex.Message));
                }
                else if (ex.Message.Contains("0x800706BA")) // Server RPC not available, occurs after a crash of OneNote
                {
                    if (!retry)
                    {
                        // 1st attempt, reinit OneNote connector and make a 2nd try

                        var delayBeforeRetrySeconds = 10;
                        LogError(page, ex, string.Format(Localizer.GetString("ErrorDuringPageProcessingRetryInProgress"), page.TitleWithPageLevelTabulation, page.Id, ex.Message, delayBeforeRetrySeconds));

                        // Recreate OneNote COM component to avoid "Server RPC not available" errors
                        OneNoteApp.CleanUp();
                        Thread.Sleep(delayBeforeRetrySeconds * 1000);
                        OneNoteApp.RenewInstance();

                        var retrySuccess = ExportPage(page, true);
                        if (retrySuccess)
                        {
                            Log.Information($"Page '{page.GetPageFileRelativePath(AppSettings.MdMaxFileLength)}': {Localizer.GetString("SuccessPageExportAfterRetry")}");
                            return true;
                        }
                        else
                            LogError(page, ex, string.Format(Localizer.GetString("ErrorDuringPageProcessing"), page.TitleWithPageLevelTabulation, page.Id, ex.Message));
                    }
                    else
                    {
                        LogError(page, ex, string.Format(Localizer.GetString("ErrorDuringPageProcessing"), page.TitleWithPageLevelTabulation, page.Id, ex.Message));
                    }
                }
                else
                {
                    LogError(page, ex, string.Format(Localizer.GetString("ErrorDuringPageProcessing"), page.TitleWithPageLevelTabulation, page.Id, ex.Message));
                }

                return false;
            }
        }

        /// <summary>
        /// Pre-process OneNote XML page for: Sections unfold, Convert OneNote tags to #hash-tags, Keep checkboxes, etc.
        /// If page XML content was changed due pre-processing, the new content stored at temporary notebook
        /// </summary>
        /// <param name="xmlPageContent">Page to pre-process</param>
        /// <returns>Temporary OneNote ID of changed pre-processed page or NULL if pre-processing do not changed page XML</returns>
        private string PageXmlPreProcessing(XElement xmlPageContent)
        {
            // Trigger on any XML tree changes so we know that this page should be cloned to temporary notebook
            var isXmlChanged = false;
            void ChangesHandler(object _, XObjectChangeEventArgs __)
            {
                isXmlChanged = true;
                xmlPageContent.Changed -= ChangesHandler;
            }
            xmlPageContent.Changed += ChangesHandler;

            var ns = xmlPageContent.Name.Namespace;

            /// Unfold page content by removing all OneNote XML attribute "collapsed" everywhere
            foreach (var xmlOutline in xmlPageContent.Descendants(ns + "OE"))
            {
                xmlOutline.Attribute("collapsed")?.Remove();
            }

           /// Keep "OneNote tag information" by adding custom tags in text content
            ConvertOnenoteTags(xmlPageContent, ns);

            /// Make indenting explicit in content by adding empty lines before text blocks
            /// NB: this has to be AFTER the ConvertOnenoteTags method, otherwise the tabs come in between the tags and the text
            if (AppSettings.IndentingStyle != IndentingStyleEnum.LeaveAsIs)
                ConvertIndentation(xmlPageContent, ns, AppSettings.IndentingStyle);

            /// Add horizontal bar before text blocks
            AddHorizontalBarBeforeTextblocks(xmlPageContent, ns);

            /// Convert hash valued colors to yellow
            if (AppSettings.convertHexValueHighlightingToYellow)
                convertHexValueHighlightingToYellow(xmlPageContent, ns);

            /// Keep HTML highlighting (using span elements). Notesnook can handle this!
            if (AppSettings.UseHtmlStyling)
            {
                // Capture font styling in span elements inside text elements
                CaptureFontStyling(xmlPageContent, ns);
                // Escape HTML span elements with style attributes, or highlighting will be removed by Pandoc
                EscapeStylingSpan(xmlPageContent, ns);
            }

            if (isXmlChanged)
                return TemporaryNotebook.ClonePage(xmlPageContent);
            else
                return null;
        }

        private static void EscapeStylingSpan(XElement xmlPageContent, XNamespace ns)
        {
            var highlightRegex = new Regex(@"<span\s+style='(\s*[a-zA-Z0-9\s\.\#;:-]*)'>(.*?)<\/span>");
            foreach (var xmlText in xmlPageContent.Descendants(ns + "T"))
            {
                if (xmlText.FirstNode is not XCData cdataNode)
                {
                    // Only log if the tag is one we expect to handle
                    Log.Warning($"Found T-element but no CDATA-element, with Value: '{xmlText?.Value}'");
                    continue;
                }
                XCData innerNode = xmlText.FirstNode as XCData;

                innerNode.Value = highlightRegex.Replace(innerNode.Value, match =>
                {
                    return $"«span style='{match.Groups[1]}'»{match.Groups[2]}«/span»";
                });
            }
        }

        private static void CaptureFontStyling(XElement xmlPageContent, XNamespace ns)
        {
            foreach (var textElement in xmlPageContent.Descendants(ns + "T"))
            {
                // Capture CDATA element
                if (textElement.FirstNode is not XCData cdataNode)
                {
                    // Only log if the tag is one we expect to handle
                    Log.Warning($"Found T-element but no CDATA-element, with Value: '{textElement?.Value}'");
                    continue;
                }
                XCData innerNode = textElement.FirstNode as XCData;

                // This is kindoff cheating, but the code is much simpler:
                //  In the case of bold/italic/underline text, the text is already inside a <span> element.
                //  However the text element itself in those cases also has a style attribute (which is basically redundant)
                //  This will capture that as well and so will put the inner span with styling, in anouther span-element with styling.
                //  However, because the escaping only captures the inner span-element, we end up with the same situation
                //      as before: just span element inside the text element.
                var styleAttribute = textElement.Attribute("style") ?? textElement.Parent?.Attribute("style");
                if (styleAttribute is not null)
                    innerNode.Value = $"<span style='{styleAttribute.Value}'>{textElement.Value}</span>";
            }
        }

        private static void convertHexValueHighlightingToYellow(XElement xmlPageContent, XNamespace ns)
        {
            // Fix for non-standard text highlights:
            // Replace OneNote CDATA HTML tags <span style="background:#SOME_HEX_VAL"> by <span style="background:yellow">
            var highlightRegex = new Regex(@"(<span\s+style='[^']*?background)\s*:\s*#\w+");
            foreach (var xmlText in xmlPageContent.Descendants(ns + "T"))
            {
                xmlText.Value = highlightRegex.Replace(xmlText.Value, match =>
                {
                    return $"{match.Groups[1]}:yellow";
                });
            }
        }

        private static readonly string HorizontalBar = "---" + Environment.NewLine + Environment.NewLine;
        private void AddHorizontalBarBeforeTextblocks(XElement xmlPageContent, XNamespace ns)
        {
            // Skip the first outline element
            foreach (var outline in xmlPageContent.Descendants(ns + "Outline").Skip(1))
            {
                // Find the first <T> element with a CDATA node (to be sure)
                var textElement = outline
                    .Descendants(ns + "T")
                    .FirstOrDefault(e => e.LastNode != null && e.LastNode.NodeType.ToString() == "CDATA");

                if (textElement == null)
                    continue;

                // Add a new line with the horizontal bar before the text element
                var emptyLineXml = new XElement(ns + "OE", new XAttribute("alignment", "left"),
                    new XElement(ns + "T", 
                    new XCData($"{HorizontalBar}")));
                textElement.Parent?.Parent?.AddFirst(emptyLineXml);
            }
        }

        const int EmSpacesPerIndent = 2;
        private void ConvertIndentation(XElement xmlPageContent, XNamespace ns, IndentingStyleEnum indentStyle)
        {
            string defaultFontSize = getQuickStyleFontsize(xmlPageContent, ns);
            foreach (var textElement in xmlPageContent.Descendants(ns + "T"))
            {
                // Determine indentation level and skip if not indented
                int indentLevel = textElement.Ancestors(ns + "OEChildren").Count() - 1;
                if (indentLevel <= 0)
                    continue;

                // If already a list, we can skip it
                var prevEl = textElement.PreviousNode as XElement;
                if (prevEl?.Name.LocalName == "List")
                    continue;

                // If inside a table, we skip it
                if (textElement.Ancestors(ns + "Table").Count() > 0)
                    continue;

                // TODO: check of in tabel!
                switch (indentStyle)
                {
                    case IndentingStyleEnum.ConvertToEmSpaces:
                        textElement.Value = Repeat("&emsp;", indentLevel * EmSpacesPerIndent) + textElement.Value;
                        break;
                    case IndentingStyleEnum.ConvertToBullets:
                        var bulletList = CreateListElement(ns, indentLevel, defaultFontSize);
                        textElement.AddBeforeSelf(bulletList);
                        break;
                }
            }
        }

        // Create list element with bullet:
        //      <one:List>
        //          <one:Bullet bullet = "13" fontSize = "11.0"/>
        //      </one:List>
        private XElement CreateListElement(XNamespace ns, int indentLevel, string fontSize)
        {
            return new XElement(ns + "List",
                new XElement(ns + "Bullet",
                    new XAttribute("bullet", indentLevel.ToString()),
                    new XAttribute("fontSize", fontSize)));
        }

        public string Repeat(string text, int n)
        {
            var textAsSpan = text.AsSpan();
            var span = new Span<char>(new char[textAsSpan.Length * n]);
            for (var i = 0; i < n; i++)
            {
                textAsSpan.CopyTo(span.Slice((int)i * textAsSpan.Length, textAsSpan.Length));
            }

            return span.ToString();
        }

        private Dictionary<string, OneNoteTagDefEnum> GetTagDefDict(XElement xmlPageContent, XNamespace ns) {
            // Get all tag definitions from the page content
            Dictionary<string, OneNoteTagDefEnum> tags = new Dictionary<string, OneNoteTagDefEnum>();

            foreach (var tagDef in xmlPageContent.Descendants(ns + "TagDef"))
            {
                if (string.IsNullOrEmpty(tagDef.Attribute("name")?.Value))
                    continue;

                var index = tagDef.Attribute("index")?.Value;
                var type = tagDef.Attribute("type")?.Value;
                var symbol = tagDef.Attribute("symbol")?.Value;
                var highlightColor = tagDef.Attribute("highlightColor")?.Value;
                var name = tagDef.Attribute("name")?.Value;
                
                // Determine tag type based on certainty of attributes
                if (type == "0" || symbol == "3" || name == "To Do" || name == "Taak" || name == "Takenlijst" || type == "99")
                    tags.Add(index, OneNoteTagDefEnum.Task);
                else if (type == "1" || symbol == "13" || name == "Important" || name == "Belangrijk" || type ==  "115")
                    tags.Add(index, OneNoteTagDefEnum.Important);
                else if (type == "2" || symbol == "15" || name == "Question" || name == "Vraag")
                    tags.Add(index, OneNoteTagDefEnum.Question);
                else if (type == "3" || name == "Remember for later" || highlightColor == "#FFFF00")
                    tags.Add(index, OneNoteTagDefEnum.RemindMeLater);
                else if (type == "4" || name == "Definition" || name == "Definitie" || type == "119" || highlightColor == "#00FF00")
                    tags.Add(index, OneNoteTagDefEnum.Definition);
            }

            return tags;
        }


        /// <summary>
        /// Convert Onenote tags to custom tags/emoticons in the text content so the tag information is conveyed to end result.
        /// In theory you could try and replace the custom tags with markdown compatible elements (e.g. for tasks), but this has too many edge cases (e.g. task in table).
        /// If you want to do this, you could use the "FinalizePageMdPostProcessing" method for this.
        /// </summary>
        /// <param name="xmlPageContent"></param>
        /// <param name="ns"></param>
        const string CustomTagUnchecked = "🔲 ";
        const string CustomTagChecked = "✅ ";
        const string CustomTagStar = "⭐ ";
        const string CustomTagQuestion = "❓ ";
        const string CustomTagRemember = "<span style='background:yellow;mso-highlight:yellow'>";
        const string CustomTagDefinition = "<span style='background:green;mso-highlight:green'>";
        private void ConvertOnenoteTags(XElement xmlPageContent, XNamespace ns)
        {
            var tags = GetTagDefDict(xmlPageContent, ns);

            // Find occurances and replace
            foreach (var tagElement in xmlPageContent.Descendants(ns + "Tag"))
            {
                // Get the corresponding text element
                XElement textElement = tagElement.Parent.Descendants(ns + "T").First() as XElement;
                if (textElement.FirstNode is not XCData)
                {
                    Log.Warning($"Found tag, but couldn't add custom tag. No CDATA-field found for element with content: '{textElement?.Value}'");
                    continue;
                }

                var tagIndex = tagElement.Attribute("index")?.Value;
                if (!tags.ContainsKey(tagIndex))
                    continue;
                                
                // Determine which custom tag to use
                var tagType = tags[tagIndex];
                string customTag;
                string highlightEndTag = "";
                if (tagType == OneNoteTagDefEnum.Task)
                    customTag = (tagElement.Attribute("completed")?.Value == "false") ? CustomTagUnchecked : CustomTagChecked;
                else if (tagType == OneNoteTagDefEnum.Important)
                    customTag = CustomTagStar;
                else if (tagType == OneNoteTagDefEnum.Question)
                    customTag = CustomTagQuestion;
                else if (tagType == OneNoteTagDefEnum.RemindMeLater)
                {
                    customTag = CustomTagRemember;
                    highlightEndTag = "</span>";
                }
                else if (tagType == OneNoteTagDefEnum.Definition)
                {
                    customTag = CustomTagDefinition;
                    highlightEndTag = "</span>";
                }
                else
                    continue; // Skip other tags

                // Add custom tag right before the tasks inner content
                XCData innerNode = textElement.FirstNode as XCData;
                innerNode.Value = $"{customTag}{innerNode.Value}{highlightEndTag}";
            }
        }

        private static string getQuickStyleFontsize(XElement xmlPageContent, XNamespace ns)
        {
            return getElementAttributeValue(xmlPageContent, ns, "QuickStyleDef", "p", "fontSize", "11.0");
        }
        private static string getElementAttributeValue(XElement xmlPageContent, XNamespace ns, string elementLabel, string elementName, string attributeLabel, string defaultValue)
        {
            return xmlPageContent
                .Descendants(ns + elementLabel)
                .FirstOrDefault(e => e.Attribute("name")?.Value == elementName)
                ?.Attribute(attributeLabel)?.Value ?? defaultValue;
        }

        protected abstract string FinalizePageMdPostProcessing(Page page, string md);

        private static void LogError(Page p, Exception ex, string message)
        {
            Log.Warning($"Page '{p.GetPageFileRelativePath(AppSettings.MdMaxFileLength)}': {message}");
            Log.Debug(ex, ex.Message);
        }

        /// <summary>
        /// Final class needs to implement logic to write the md file of the page in the export folder
        /// </summary>
        /// <param name="page">The page</param>
        /// <param name="pageMd">Markdown content of the page</param>
        protected abstract void WritePageMdFile(Page page, string pageMd);


        /// <summary>
        /// Create attachment files in export folder, and update page's markdown to insert md reference that link to the attachment files
        /// </summary>
        /// <param name="page"></param>
        /// <param name="pageMdFileContent">Markdown content of the page</param>
        private void ExportPageAttachments(Page page, ref string pageMdFileContent)
        {
            foreach (Attachement attach in page.Attachements)
            {
                if (attach.Type == AttachementType.File)
                {
                    EnsureAttachmentFileIsNotUsed(page, attach);

                    var exportFilePath = GetAttachmentFilePath(attach);

                    Directory.CreateDirectory(Path.GetDirectoryName(exportFilePath));

                    // Copy attachment file into export folder
                    File.Copy(attach.ActualSourceFilePath, exportFilePath);
                    //File.SetAttributes(exportFilePath, FileAttributes.Normal); // Prevent exception during removing of export directory

                    // Update page markdown to insert md references to attachments
                    InsertPageMdAttachmentReference(ref pageMdFileContent, attach, GetAttachmentMdReference);
                }

                FinalizeExportPageAttachments(page, attach);
            }
        }


        /// <summary>
        /// Final class needs to implement logic to write the md file of the attachment file in the export folder (if needed)
        /// </summary>
        /// <param name="page">The page</param>
        /// <param name="attachment">The attachment</param>
        protected abstract void FinalizeExportPageAttachments(Page page, Attachement attachment);


        /// <summary>
        /// Replace the tag <<FileName>> generated by OneNote by a markdown link referencing the attachment
        /// </summary>
        /// <param name="pageMdFileContent"></param>
        /// <param name="attach"></param>
        private static void InsertPageMdAttachmentReference(ref string pageMdFileContent, Attachement attach, Func<Attachement, string> getAttachMdReferenceMethod)
        {
            var pageMdFileContentModified = Regex.Replace(pageMdFileContent, "(\\\\<){2}(?<fileName>.*)(\\\\>){2}", delegate (Match match)
            {
                var refFileName = match.Groups["fileName"]?.Value ?? "";
                var attachOriginalFileName = attach.OneNotePreferredFileName;
                var attachMdRef = getAttachMdReferenceMethod(attach);

                if (refFileName.Equals(attachOriginalFileName))
                {
                    // reference found is corresponding to the attachment being processed
                    return $"[{attachOriginalFileName}]({attachMdRef})";
                }
                else
                {
                    // not the current attachment, ignore
                    return match.Value;
                }
            });

            pageMdFileContent = pageMdFileContentModified;
        }


        /// <summary>
        /// Replace PanDoc IMG HTML tag by markdown reference and copy image file into notebook export directory
        /// </summary>
        /// <param name="page">Section page</param>
        /// <param name="mdFileContent">Content of the MD file</param>
        /// <param name="resourceFolderPath">The path to the notebook folder where store attachments</param>
        public void ExtractImagesToResourceFolder(Page page, ref string mdFileContent)
        {
            // Replace <IMG> tags by markdown references
            var pageTxtModified = Regex.Replace(mdFileContent, "<img [^>]+/>", delegate (Match match)
            {
                string imageTag = match.ToString();

                // http://regexstorm.net/tester
                string regexImgAttributes = "<img src=\"(?<src>[^\"]+)\".* />";

                MatchCollection matches = Regex.Matches(imageTag, regexImgAttributes, RegexOptions.IgnoreCase);
                Match imgMatch = matches[0];

                var panDocHtmlImgTagPath = Path.GetFullPath(imgMatch.Groups["src"].Value);
                panDocHtmlImgTagPath = WebUtility.HtmlDecode(panDocHtmlImgTagPath);
                Attachement imgAttach = page.ImageAttachements.Where(img => PathExtensions.PathEquals(img.ActualSourceFilePath, panDocHtmlImgTagPath)).FirstOrDefault();

                // Only add a new attachment if this is the first time the image is referenced in the page
                if (imgAttach == null)
                {
                    // Add a new attachment to current page
                    imgAttach = new Attachement(page)
                    {
                        Type = AttachementType.Image,
                        ActualSourceFilePath = Path.GetFullPath(panDocHtmlImgTagPath),
                        OriginalUserFilePath = Path.GetFullPath(panDocHtmlImgTagPath) // Not really a user file path but a PanDoc temp file
                    };

                    page.Attachements.Add(imgAttach);

                    EnsureAttachmentFileIsNotUsed(page, imgAttach);
                }

                var attachRef = GetAttachmentMdReference(imgAttach);
                var refLabel = Path.GetFileNameWithoutExtension(imgAttach.ActualSourceFilePath);
                return $"![{refLabel}]({attachRef})";

            });


            // Move attachments file into output resource folder and delete tmp file
            // In case of duplicate files, suffix attachment file name
            foreach (var attach in page.ImageAttachements)
            {
                var attachFilePath = GetAttachmentFilePath(attach);
                Directory.CreateDirectory(Path.GetDirectoryName(attachFilePath));
                File.Copy(attach.ActualSourceFilePath, attachFilePath);
                File.Delete(attach.ActualSourceFilePath);
            }


            if (AppSettings.PostProcessingMdImgRef)
            {
                mdFileContent = pageTxtModified;
            }
        }

        /// <summary>
        /// Suffix the attachment file name if it conflicts with an other attachment previously attached to the notebook export
        /// </summary>
        /// <param name="page">The parent Page</param>
        /// <param name="attach">The attachment</param>
        private void EnsureAttachmentFileIsNotUsed(Page page, Attachement attach)
        {
            var notUseFileNameFound = false;
            var cmpt = 0;
            var attachmentFilePath = GetAttachmentFilePath(attach);

            while (!notUseFileNameFound)
            {
                var candidateFilePath = cmpt == 0 ? attachmentFilePath :
                    $"{Path.ChangeExtension(attachmentFilePath, null)}-{cmpt}{Path.GetExtension(attachmentFilePath)}";

                var attachmentFileNameAlreadyUsed = page.GetNotebook().GetAllAttachments().Any(a => a != attach && PathExtensions.PathEquals(GetAttachmentFilePath(a), candidateFilePath));

                // because of using guid, this step should no longer needed and need to be removed
                if (!attachmentFileNameAlreadyUsed)
                {
                    if (cmpt > 0)
                        attach.OverrideExportFilePath = candidateFilePath;

                    notUseFileNameFound = true;
                }
                else
                    cmpt++;
            }

        }


        /// <summary>
        /// Suffix the page file name if it conflicts with an other page previously attached to the notebook export
        /// </summary>
        /// <param name="page">The parent Page</param>
        /// <param name="attach">The attachment</param>
        private void EnsurePageUniquenessPerSection(Page page)
        {
            var notUseFileNameFound = false;
            var cmpt = 0;
            var pageFilePath = GetPageMdFilePath(page);

            while (!notUseFileNameFound)
            {
                var candidateFilePath = cmpt == 0 ? pageFilePath :
                    $"{Path.ChangeExtension(pageFilePath, null)}-{cmpt}.md";

                var attachmentFileNameAlreadyUsed = page.Parent.Childs.OfType<Page>().Any(p => p != page && PathExtensions.PathEquals(GetPageMdFilePath(p), candidateFilePath));

                if (!attachmentFileNameAlreadyUsed)
                {
                    if (cmpt > 0)
                        page.OverridePageFilePath = candidateFilePath;

                    notUseFileNameFound = true;
                }
                else
                    cmpt++;
            }
        }

        private static void ProcessPageAttachments(XNamespace ns, Page page, XElement xmlPageContent)
        {
            foreach (var xmlAttachment in xmlPageContent.Descendants(ns + "InsertedFile").Concat(xmlPageContent.Descendants(ns + "MediaFile")))
            {
                var fileAttachment = new Attachement(page)
                {
                    ActualSourceFilePath = xmlAttachment.Attribute("pathCache")?.Value,
                    OriginalUserFilePath = xmlAttachment.Attribute("pathSource")?.Value,
                    OneNotePreferredFileName = xmlAttachment.Attribute("preferredName")?.Value,
                    Type = AttachementType.File
                };

                if (fileAttachment.ActualSourceFilePath != null)
                {
                    page.Attachements.Add(fileAttachment);
                }
            }
        }
    }
}
