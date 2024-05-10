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
using Microsoft.AspNetCore.Html;

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
                "dczpqymrv",                    // Cloud name
                "545529419662769",              // API Key
                "U6CSGR8_K6_WMr3yEpxBO8T2Ka4"   // API Secret
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
                    // ************* Upload file word *************
                    if (extension == ".doc" || extension == ".docx")
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            file.CopyTo(memoryStream);
                            string fileName = Path.GetFileNameWithoutExtension(file.FileName);
                            bool created = await ConvertToHtml(memoryStream, fileName, categories, url); // Convert word sang html
                            if (created) result.Success.Add(file.FileName);
                            else result.Failed.Add(file.FileName);
                        }
                    }
                    // ************* Upload file zip *************
                    else if (extension == ".zip")
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            file.CopyTo(memoryStream);
                            memoryStream.Position = 0;

                            // ************* Giải nén file zip *************
                            using (var archive = new ZipArchive(memoryStream))
                            {
                                if (archive.Entries.Any(entry => !IsDocOrDocx(entry.FullName))) // Return nếu zip chứa file không đúng định dạng
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

                                        bool created = await ConvertToHtml(entryMemoryStream, fileName, categories, url); // Convert word sang html
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

                if (result.Success.Count > 0) return new ResponseResult<CreatePost>(RetCodeEnum.Ok, "Upload file thành công", result);
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

        private async Task<bool> ConvertToHtml(Stream memoryStream, string fileName, string category, string url)
        {
            bool result = false;
            int index = 1;
            memoryStream.Position = 0;
            string htmlString = "";
            string thumbnail = "";

            // Setting convert html
            HtmlConverterSettings convSettings = new HtmlConverterSettings()
            {
                FabricateCssClasses = true,                       // Tự động generate các class css
                CssClassPrefix = "cls-",                          // Tiền tố css
                RestrictToSupportedLanguages = false,             // Tất cả các ngôn ngữ sẽ được chuyển đổi, ngay cả khi không được hỗ trợ.
                RestrictToSupportedNumberingFormats = false,      // Tất cả các định dạng số sẽ được chuyển đổi, ngay cả khi không được hỗ trợ.
                AdditionalCss = "span { width: fit-content!important } img { margin : 6px 0px; height: auto!important;}",
                ImageHandler = (imageInfo) =>                     // Config hình ảnh
                {
                    async Task<XElement> HandleImageAsync(ImageInfo info)
                    {
                        string mediaType = info.ContentType;
                        CloudResult cloudResult = await Upload(imageInfo, "word_images", mediaType);
                        string imagePath = cloudResult != null && cloudResult.status == 200 ? cloudResult.path : "";
                        if (index == 1)
                        {
                            thumbnail = imagePath; // thumbnail là ảnh đầu tiên trong file
                        }

                        // Chuyển đổi hình ảnh trong file word
                        XElement img = new XElement(Xhtml.img,
                                        new XAttribute(NoNamespace.src, imagePath),
                                        imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null,
                                        new XAttribute(NoNamespace.style, "width: auto; height: auto;"));

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
                    // ***************** Format numbering alignment thành Left ****************
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

                    // ***************** Tiến hành convert ****************
                    XElement html = OpenXmlPowerTools.HtmlConverter.ConvertToHtml(doc, convSettings);

                    XElement htmlElement = new XElement(Xhtml.html, html);

                    htmlString = htmlElement.ToString();
                    htmlString = RemoveLRMCharacter(htmlString); // Remove các kí tự đặc biệt
                    htmlString = RemoveHtmlTags(htmlString);     // Remove các thẻ <html/> , <meta/> , <head/>
                    htmlString = await InlinerCSS(htmlString);   // Css Inline
                    
                    htmlString = Regex.Replace(htmlString, "\\n", "");
                    htmlString = Regex.Replace(htmlString, "\\r", "");
                    htmlString = RemoveFontFamily(htmlString);
                    string pattern = @"font-size\s*:\s*[^;']*;";
                    htmlString = Regex.Replace(htmlString, pattern, string.Empty);
                    if (!string.IsNullOrEmpty(htmlString))
                    {
                        result = await _post.AddPost(htmlString, fileName, thumbnail, category, url); // Tạo post
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
        private async Task<CloudResult> Upload(ImageInfo image, string folder, string type)
        {
            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    if (type.Contains("gif")) image.Bitmap.Save(memoryStream, ImageFormat.Gif);
                    else image.Bitmap.Save(memoryStream, ImageFormat.Jpeg);
                    memoryStream.Position = 0;
                    byte[] imageBytes = memoryStream.ToArray();

                    string base64String = Convert.ToBase64String(imageBytes);
                    ImageUploadParams uploadParams = new ImageUploadParams
                    {
                        File = new FileDescription(new Guid().ToString(), memoryStream),
                        PublicId = Guid.NewGuid().ToString(),
                        Folder = folder,
                        Transformation = null
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

        [HttpGet("test")]
        public IActionResult test(string html)
        {
           
            return Ok(RemoveFontFamily(html));
        }

        private string RemoveFontFamily(string html)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);

            HtmlNodeCollection spanNodes = doc.DocumentNode.SelectNodes("//span");

            if (spanNodes != null)
            {
                foreach (HtmlNode node in spanNodes)
                {
                    string fontFamily = GetFontFamily(node.GetAttributeValue("style", ""));

                    if (fontFamily != "Symbol")
                    {
                        node.Attributes["style"].Value = HandleRemoveFontFamily(node.GetAttributeValue("style", ""));
                    }
                }
            }


            HtmlNodeCollection pNode = doc.DocumentNode.SelectNodes("//p");
            if (pNode != null)
            {
                foreach (HtmlNode node in pNode)
                {
                    node.Attributes["style"].Value = HandleRemoveFontFamily(node.GetAttributeValue("style", ""));
                }
            }

            string modifiedHtml = doc.DocumentNode.OuterHtml;
            return modifiedHtml;
        }

        private string GetFontFamily(string style)
        {
            string[] attributes = style.Split(';');
            foreach (string attribute in attributes)
            {
                if (attribute.Trim().StartsWith("font-family:"))
                {
                    string fontFamily = attribute.Trim().Substring(12).Trim();
                    return fontFamily;
                }
            }
            return "";
        }

        private string HandleRemoveFontFamily(string style)
        {
            string[] attributes = style.Split(';');
            List<string> updatedAttributes = new List<string>();
            foreach (string attribute in attributes)
            {
                if (!attribute.Trim().StartsWith("font-family"))
                {
                    updatedAttributes.Add(attribute.Trim());
                }
            }
            return string.Join("; ", updatedAttributes);
        }
    }
}
