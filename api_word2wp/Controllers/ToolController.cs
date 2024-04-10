using api_word2wp.Interfaces;
using api_word2wp.Models;
using CloudinaryDotNet.Actions;
using CloudinaryDotNet;
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OpenXmlPowerTools;
using RestSharp;
using System.Collections.Generic;
using System.IO.Compression;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System;
using api_word2wp.Response;
using ImageInfo = OpenXmlPowerTools.ImageInfo;
using System.Drawing.Imaging;
using System.Linq;
using DocumentFormat.OpenXml.Vml.Office;
using System.Text.RegularExpressions;
using System.Net.Http;
using DocumentFormat.OpenXml.Wordprocessing;

namespace api_word2wp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ToolController : ControllerBase
    {
        private readonly Cloudinary _cloudinary;
        private readonly ICategoryService _category;
        private readonly IPostService _post;
        public ToolController(ICategoryService category, IPostService post)
        {
            var cloudinaryAccount = new Account(
                "dczpqymrv",
                "545529419662769",
                "U6CSGR8_K6_WMr3yEpxBO8T2Ka4"
            );
            _cloudinary = new Cloudinary(cloudinaryAccount);
            _category = category;
            _post = post;
        }

        [HttpPost("upload")]
        public async Task<ResponseResult<CreatePost>> UploadFileZip(List<IFormFile> files, string categories, string url)
        {
            try
            {
                if (files.Count == 0) return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "Vui lòng chọn file", new CreatePost());
                if (string.IsNullOrEmpty(categories)) return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "Vui lòng chọn thể loại ", new CreatePost());
                if (string.IsNullOrEmpty(url)) return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "Vui lòng nhập địa chỉ trang web ", new CreatePost());
                CreatePost result = new CreatePost();
                foreach (IFormFile file in files)
                {
                    var extension = System.IO.Path.GetExtension(file.FileName);
                    if (extension == ".doc" || extension == ".docx")
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            file.CopyTo(memoryStream);
                            string fileName = Path.GetFileNameWithoutExtension(file.FileName);
                            bool created = await ConvertToHtml(memoryStream, fileName, categories, url);
                            if (created) result.Success.Add(file.FileName);
                            else result.Failed.Add(file.FileName);
                        }
                    }
                    else if (extension == ".zip")
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            file.CopyTo(memoryStream);
                            memoryStream.Position = 0;

                            using (var archive = new ZipArchive(memoryStream))
                            {
                                if (archive.Entries.Any(entry => !IsDocOrDocx(entry.FullName)))
                                {
                                    return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "File zip chỉ có thể chứa file doc hoặc docx", new CreatePost());
                                }
                                foreach (var entry in archive.Entries)
                                {
                                    var entryMemoryStream = new MemoryStream();
                                    string fileName = Path.GetFileNameWithoutExtension(entry.Name);
                                    using (var entryStream = entry.Open())
                                    {
                                        entryStream.CopyTo(entryMemoryStream);
                                        entryMemoryStream.Position = 0;

                                        bool created = await ConvertToHtml(entryMemoryStream, fileName, categories, url);
                                        if (created) result.Success.Add(entry.FullName);
                                        else result.Failed.Add(entry.FullName);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "File không đúng định dạng", new CreatePost());
                    }
                }

                if(result.Success.Count > 0) return new ResponseResult<CreatePost>(RetCodeEnum.Ok, "Upload file thành công", result);
                else return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "Upload file thất bại", result);

            }
            catch (Exception ex)
            {
                return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "Đã có lỗi xảy ra", new CreatePost());
            }
        }

        [HttpGet("categories")]
        public async Task<ResponseResult<List<WpCategory>>> List(string url)
        {
            List<WpCategory> categories = await _category.GetList(url);
            return new ResponseResult<List<WpCategory>>(RetCodeEnum.Ok, RetCodeEnum.Ok.ToString(), categories);
        }

        private async Task<bool> ConvertToHtml(Stream memoryStream, string fileName, string category,string url)
        {
            bool result = false;
            int index = 1;
            memoryStream.Position = 0;
            string htmlString = "";
            string thumbnail = "";

            HtmlConverterSettings convSettings = new HtmlConverterSettings()
            {
                FabricateCssClasses = true,
                CssClassPrefix = "cls-",
                RestrictToSupportedLanguages = false,
                RestrictToSupportedNumberingFormats = false,
                AdditionalCss = "span { width: fit-content!important } img { margin : 6px 0px; } body { font-family: Arial, sans-serif !important;}",
                ImageHandler = (imageInfo) =>
                {
                    async Task<XElement> HandleImageAsync(ImageInfo info)
                    {
                        CloudResult cloudResult = await Upload(imageInfo, "word_images");
                        string imagePath = cloudResult != null && cloudResult.status == 200 ? cloudResult.path : "";
                        if (index == 1)
                        {
                            thumbnail = imagePath;
                        }
                        XElement img = new XElement(Xhtml.img,
                               new XAttribute(NoNamespace.src, imagePath),
                               imageInfo.ImgStyleAttribute,
                               imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);

                        index++;
                        return img;

                    }
                    return HandleImageAsync(imageInfo).Result;
                }
            };

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var numberingDefinitionsPart = doc.MainDocumentPart.NumberingDefinitionsPart;
                    if (numberingDefinitionsPart != null)
                    {
                        Numbering numbering = numberingDefinitionsPart.Numbering;

                        if (numbering != null)
                        {
                            foreach (var num in numbering.Descendants<NumberingInstance>())
                            {
                                string abstractNumId = num.AbstractNumId.Val;
                                AbstractNum abstractNum = numbering.Descendants<AbstractNum>().FirstOrDefault(an => an.AbstractNumberId == abstractNumId);

                                if (abstractNum != null)
                                {
                                    foreach (var level in abstractNum.Descendants<Level>())
                                    {
                                        level.LevelJustification = new LevelJustification { Val = LevelJustificationValues.Left };
                                    }
                                }
                            }
                            doc.Save();
                        }
                    }

                    XElement html = OpenXmlPowerTools.HtmlConverter.ConvertToHtml(doc, convSettings);

                    XElement htmlElement = new XElement(Xhtml.html, html);

                    htmlString = htmlElement.ToString();
                    htmlString = RemoveLRMCharacter(htmlString);
                    htmlString = RemoveHtmlTags(htmlString);
                    htmlString = await InlinerCSS(htmlString);

                    if (!string.IsNullOrEmpty(htmlString))
                    {
                        result = await _post.AddPost(htmlString, fileName, thumbnail, category, url);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return result;
        }

        #region ======= Xử lí html =======
        private string RemoveLRMCharacter(string htmlString)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(htmlString);

            foreach (var textNode in doc.DocumentNode.SelectNodes("//text()"))
            {
                var cleanedText = Regex.Replace(textNode.InnerText, @"[\u200E\u200F]", "");
                textNode.InnerHtml = HtmlEntity.DeEntitize(cleanedText);
            }

            return doc.DocumentNode.OuterHtml;
        }
        private string RemoveHtmlTags(string htmlString)
        {
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlString);
            HtmlNode headNode = htmlDoc.DocumentNode.SelectSingleNode("//head");
            HtmlNode styleNode = headNode.SelectSingleNode("style");
            string styleContent = styleNode.InnerHtml;
            StringBuilder newHtmlBuilder = new StringBuilder();
            newHtmlBuilder.Append("<style>");
            newHtmlBuilder.Append(styleContent);
            newHtmlBuilder.Append("</style>");
            newHtmlBuilder.Append(htmlDoc.DocumentNode.SelectSingleNode("//body").OuterHtml);
            return newHtmlBuilder.ToString();
        }
        private async Task<string> InlinerCSS(string html)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var formData = new MultipartFormDataContent();
                    formData.Add(new StringContent(html), "html");
                    var response = await client.PostAsync("https://templates.mailchimp.com/services/inline-css/", formData);
                    if (response.IsSuccessStatusCode)
                    {
                        var content = await response.Content.ReadAsStringAsync();
                        if (!content.Contains("Invalid Content"))
                        {
                            return System.Web.HttpUtility.HtmlDecode(content);
                        }
                    }

                    return html;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("inliner css failed");
                return html;
            }
        }
        private bool IsDocOrDocx(string fileName)
        {
            return fileName.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                   fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase);
        }
        private async Task<CloudResult> Upload(ImageInfo image, string folder)
        {
            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    image.Bitmap.Save(memoryStream, ImageFormat.Jpeg);
                    memoryStream.Position = 0;
                    byte[] imageBytes = memoryStream.ToArray();

                    string base64String = Convert.ToBase64String(imageBytes);
                    ImageUploadParams uploadParams = new ImageUploadParams
                    {
                        File = new FileDescription(new Guid().ToString(), memoryStream),
                        PublicId = Guid.NewGuid().ToString(),
                        Folder = folder,
                    };

                    ImageUploadResult result = await _cloudinary.UploadAsync(uploadParams);

                    if (result.Error != null)
                    {
                        return new CloudResult();
                    }

                    return new CloudResult(result.DisplayName, result.SecureUrl.ToString(), result.PublicId, (int)result.StatusCode);
                }
            }
            catch (Exception ex)
            {
                return new CloudResult();
            }

        }
        private class CloudResult
        {
            public CloudResult()
            {
                name = "";
                path = "";
                publicId = "";
                status = 0;
            }

            public CloudResult(string name, string path, string publicId, int status)
            {
                this.name = name;
                this.path = path;
                this.publicId = publicId;
                this.status = status;
            }

            public string name { get; set; } = "";
            public string path { get; set; } = "";
            public string publicId { get; set; } = "";
            public int status { get; set; } = 0;
        }
        #endregion
        private void ConvertNumberingFromRomanToDecimal(WordprocessingDocument doc)
        {
            try
            {
                var numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;
                if (numberingPart != null)
                {
                    var numbering = numberingPart.Numbering;
                    foreach (var abstractNum in numbering.Elements<AbstractNum>())
                    {
                        foreach (var lvl in abstractNum.Elements<Level>())
                        {
                            if (lvl.NumberingFormat != null && lvl.NumberingFormat.Val.HasValue && lvl.NumberingFormat.Val.Value == NumberFormatValues.LowerRoman)
                            {
                                lvl.NumberingFormat.Val = NumberFormatValues.Decimal;
                            }
                        }
                    }
                    numbering.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi trong quá trình chuyển đổi đánh số: {ex.Message}");
            }
        }
    }
}
