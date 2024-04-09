using api_word2wp.Interfaces;
using System.Threading.Tasks;
using System;
using RestSharp;
using System.Net.Http;

namespace api_word2wp.Implements
{
    public class PostService : IPostService
    {
        public PostService() { }
        public async Task<bool> AddPost(string content, string title, string thumbnail, string categories, string url)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    string api = $"{url}/wp-json/mbwsapi/v1/create-post/";

                    // Tạo FormData
                    var formData = new MultipartFormDataContent();
                    formData.Add(new StringContent(title), "title");
                    formData.Add(new StringContent(content), "content");
                    formData.Add(new StringContent(categories), "categories");
                    formData.Add(new StringContent(thumbnail), "thumb");
                    formData.Add(new StringContent("publish"), "status");

                    // Thêm header Api-Key
                    client.DefaultRequestHeaders.Add("Api-Key", "MatBaoWS@1234");

                    // Gửi yêu cầu POST
                    var response = await client.PostAsync(api, formData);

                    // Xử lý kết quả trả về
                    if (response.IsSuccessStatusCode)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        /*public async Task<string> UploadImageContent(MemoryStream imageStream, string filename, string token)
        {
            try
            {
                byte[] imageData = imageStream.ToArray();

                var client = new RestClient(_wordpress.Url);
                var request = new RestRequest("/wp-json/wp/v2/media", Method.Post);
                request.AddHeader("Authorization", $"Bearer {token}");
                request.AddHeader("Content-Disposition", $"from-data; filename={filename}");
                request.AddHeader("Content-Type", "image/png"); 
                request.AddParameter("application/octet-stream", imageData, ParameterType.RequestBody);

                var response = await client.ExecutePostAsync(request);

                // Kiểm tra phản hồi từ WordPress
                if (response.StatusCode == HttpStatusCode.Created)
                {
                    // Phân tích phản hồi để lấy đường dẫn của hình ảnh
                    var jsonResponse = JObject.Parse(response.Content);
                    string imageUrl = jsonResponse.Value<string>("source_url");
                    return imageUrl;
                }
                else
                {
                    // Trả về null nếu có lỗi khi tải lên hình ảnh
                    return null;
                }
            }
            catch (Exception ex)
            {
                // Trả về null nếu có lỗi xảy ra
                return null;
            }
        }*/

    }
}
