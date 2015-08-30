using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.XPath;
using HtmlAgilityPack;
using OutlookSaveHtmlWithImages.Properties;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OutlookSaveHtmlWithImages
{
    static class HtmlSaver
    {
        private enum ModifyStatus
        {
            Unknown,
            Ignored,
            Success,
            Failure
        }

        private static ModifyStatus ModifyLink(HtmlNode element, string attribName, string outputDir, string filesDirName)
        {
            string attribValue = element.GetAttributeValue(attribName, null);
            if (attribValue == null)
            {
                return ModifyStatus.Ignored;
            }

            string attribValueLower = attribValue.ToLowerInvariant();
            if (attribValueLower.StartsWith("http://") || attribValueLower.StartsWith("https://"))
            {
                // downloadable!
                var downloadRequest = (HttpWebRequest)WebRequest.Create(attribValue);
                downloadRequest.Timeout = 5000;

                WebResponse downloadResponse;
                try
                {
                    downloadResponse = downloadRequest.GetResponse();
                }
                catch (WebException)
                {
                    // this failed; continue with the next image
                    return ModifyStatus.Failure;
                }

                byte[] imageData;
                string mimeType;

                using (downloadResponse)
                using (var responseHolder = new MemoryStream())
                {
                    downloadResponse.GetResponseStream().CopyTo(responseHolder);
                    mimeType = downloadResponse.Headers[HttpResponseHeader.ContentType] ?? "application/octet-stream";
                    imageData = responseHolder.ToArray();
                }

                string base64ImageData = Convert.ToBase64String(imageData);
                var imgDataUri = $"data:{mimeType};base64,{base64ImageData}";

                element.SetAttributeValue(attribName, imgDataUri);
                return ModifyStatus.Success;
            }

            if (attribValueLower.StartsWith(filesDirName))
            {
                byte[] fileData;
                string mimeType = "application/octet-stream";

                string inputFileName = Path.Combine(outputDir, attribValue.Replace('/', Path.DirectorySeparatorChar));
                using (var inputFile = new FileStream(inputFileName, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var contentHolder = new MemoryStream())
                {
                    inputFile.CopyTo(contentHolder);
                    fileData = contentHolder.ToArray();
                }

                string base64FileData = Convert.ToBase64String(fileData);
                var fileDataUri = $"data:{mimeType};base64,{base64FileData}";

                element.SetAttributeValue(attribName, fileDataUri);

                // try to delete original
                File.Delete(inputFileName);

                return ModifyStatus.Success;
            }

            return ModifyStatus.Ignored;
        }

        public static void PostprocessHtml(string fileName)
        {
            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(fileName);

            string outputDir = Path.GetDirectoryName(fileName);
            string filesDirName = Path.GetFileNameWithoutExtension(fileName) + "_files";

            int successCount = 0, failureCount = 0;
            var elementsToAttributes = new Dictionary<string, string>
            {
                { "img", "src" },
                { "link", "href" }
            };
            foreach (var elementToAttribute in elementsToAttributes)
            {
                string xPathString = $".//{elementToAttribute.Key}[@{elementToAttribute.Value}]";
                foreach (HtmlNode linkyElement in htmlDoc.DocumentNode.SelectNodes(xPathString))
                {
                    ModifyStatus status;
                    try
                    {
                        status = ModifyLink(linkyElement, elementToAttribute.Value, outputDir, filesDirName);
                    }
                    catch (Exception)
                    {
                        status = ModifyStatus.Failure;
                    }

                    switch (status)
                    {
                        case ModifyStatus.Success:
                            ++successCount;
                            break;
                        case ModifyStatus.Failure:
                            ++failureCount;
                            break;
                    }
                }
            }

            // overwrite
            htmlDoc.Save(fileName);

            var failurePiece = (failureCount == 0)
                ? string.Format(Resources.ExportCompletedMessagePieceNoFailures)
                : string.Format(Resources.ExportCompletedMessagePieceMoreThanOneFailure, failureCount);
            var exportCompletedMessage = string.Format(Resources.ExportCompletedMessage, successCount, failurePiece);

            MessageBox.Show(exportCompletedMessage, Resources.ExportCompletedTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
