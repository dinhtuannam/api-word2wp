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

        private async Task<bool> ConvertToHtml(Stream memoryStream, string fileName, string category)
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
                AdditionalCss = ".cls-000016 { width: 100%!important }",
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
                    XElement html = OpenXmlPowerTools.HtmlConverter.ConvertToHtml(doc, convSettings);
                    
                    XElement htmlElement = new XElement(Xhtml.html, html);

                    htmlString = htmlElement.ToString().Replace("&#x200f;", "");

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
                    htmlString = newHtmlBuilder.ToString();
                    htmlString = "<div id=\"wp-container\">" + htmlString + "</div>";
                    if (!string.IsNullOrEmpty(htmlString))
                    {
                        htmlString = await InlinerCSS(htmlString);
                        result = await _post.AddPost(htmlString, fileName, thumbnail, category);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return result;
        }


        [HttpPost("upload")]
        public async Task<ResponseResult<CreatePost>> UploadFileZip(IFormFile file, string categories)
        {
            try
            {
                if (string.IsNullOrEmpty(categories)) return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "Vui lòng chọn thể loại ", new CreatePost());
                var extension = System.IO.Path.GetExtension(file.FileName);
                CreatePost result = new CreatePost();
                if (extension == ".doc" || extension == ".docx")
                {
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        file.CopyTo(memoryStream);
                        string fileName = Path.GetFileNameWithoutExtension(file.FileName);
                        bool created = await ConvertToHtml(memoryStream, fileName, categories);
                        if (created) result.Success.Add(file.FileName);
                        else result.Failed.Add(file.FileName);
                        return new ResponseResult<CreatePost>(RetCodeEnum.Ok, RetCodeEnum.Ok.ToString(), result);
                    }

                }
                else if (extension == ".zip")
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        file.CopyTo(memoryStream);
                        memoryStream.Position = 0;

                        // Giải nén tệp ZIP từ MemoryStream
                        using (var archive = new ZipArchive(memoryStream))
                        {
                            // Kiểm tra nếu có bất kỳ tệp nào không đúng định dạng
                            if (archive.Entries.Any(entry => !IsDocOrDocx(entry.FullName)))
                            {
                                return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "Tồn tại file không đúng định dạng trong zip", new CreatePost());
                            }

                            foreach (var entry in archive.Entries)
                            {
                                var entryMemoryStream = new MemoryStream();
                                string fileName = Path.GetFileNameWithoutExtension(entry.Name);
                                using (var entryStream = entry.Open())
                                {
                                    entryStream.CopyTo(entryMemoryStream);
                                    entryMemoryStream.Position = 0;

                                    bool created = await ConvertToHtml(entryMemoryStream, fileName, categories);
                                    if (created) result.Success.Add(entry.FullName);
                                    else result.Failed.Add(entry.FullName);
                                }

                            }
                        }
                        return new ResponseResult<CreatePost>(RetCodeEnum.Ok, RetCodeEnum.Ok.ToString(), result);
                    }
                    /*var rootName = file.FileName;
                    var tempFolderPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName());
                    Directory.CreateDirectory(tempFolderPath);

                    var zipFilePath = System.IO.Path.Combine(tempFolderPath, file.FileName);
                    using (var fileStream = new FileStream(zipFilePath, FileMode.Create))
                    {
                        await file.CopyToAsync(fileStream);
                    }

                    ZipFile.ExtractToDirectory(zipFilePath, tempFolderPath);

                    var invalidFiles = Directory.EnumerateFiles(tempFolderPath, "*.*", SearchOption.AllDirectories)
                                   .Where(filePath => !IsDocOrDocx(filePath, rootName))
                                   .ToList();


                    if (invalidFiles.Any())
                    {
                        return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "File không đúng định dạng", new CreatePost());
                    }

                    foreach (var docxFile in Directory.GetFiles(tempFolderPath, "*.doc*", SearchOption.AllDirectories))
                    {
                        using (MemoryStream docxMemoryStream = new MemoryStream())
                        {
                            using (FileStream docxFileStream = new FileStream(docxFile, FileMode.Open, FileAccess.Read))
                            {
                                await docxFileStream.CopyToAsync(docxMemoryStream);
                            }
                            string fileName = System.IO.Path.GetFileNameWithoutExtension(docxFile);
                            bool created = await ConvertToHtml(docxMemoryStream, fileName, categories);
                            if (created) result.Success.Add(fileName);
                            else result.Failed.Add(fileName);
                        }
                    }

                    return new ResponseResult<CreatePost>(RetCodeEnum.Ok, RetCodeEnum.Ok.ToString(), result);*/

                }
                else
                {
                    return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "File không đúng định dạng", new CreatePost());
                }
            }
            catch (Exception ex)
            {
                return new ResponseResult<CreatePost>(RetCodeEnum.ApiError, "Đã có lỗi xảy ra", new CreatePost());
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


        [HttpGet("categories")]
        public async Task<ResponseResult<List<WpCategory>>> List()
        {
            List<WpCategory> categories = await _category.GetList();
            return new ResponseResult<List<WpCategory>>(RetCodeEnum.Ok, RetCodeEnum.Ok.ToString(), categories);
        }

        private async Task<string> InlinerCSS(string html)
        {
            try
            {
                var client = new RestClient("https://templates.mailchimp.com/services/inline-css/");
                var request = new RestRequest();
                request.AddHeader("content-type", "application/x-www-form-urlencoded");
                request.AddParameter("application/x-www-form-urlencoded", $"html={html}", ParameterType.RequestBody);
                var response = await client.ExecutePostAsync(request);
                if (response.StatusCode == System.Net.HttpStatusCode.OK && !response.Content.Contains("Invalid Content"))
                {
                    return response.Content;
                }
                return html;
            }
            catch (Exception ex)
            {
                Console.WriteLine("inliner css failed");
                return html;
            }
        }
    }
}
