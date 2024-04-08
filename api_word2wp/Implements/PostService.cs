using api_word2wp.Interfaces;
using System.Threading.Tasks;
using System;
using RestSharp;

namespace api_word2wp.Implements
{
    public class PostService : IPostService
    {
        public PostService() { }
        public async Task<bool> AddPost(string content, string title, string thumbnail, string categories)
        {
            try
            {
                string url = "https://development.matbao.website/wp-json/mbwsapi/v1/create-post/";
                var client = new RestClient(url);
                var request = new RestRequest();
                request.AddHeader("Api-Key", "MatBaoWS@1234");
                request.AddHeader("content-type", "application/x-www-form-urlencoded");
                request.AddParameter("application/x-www-form-urlencoded", $"title={title}&content={content}&categories={categories}&status=publish&thumb={thumbnail}", ParameterType.RequestBody);
                var response = await client.ExecutePostAsync(request);
                if (response.StatusCode != System.Net.HttpStatusCode.OK)
                {
                    return false;
                }
                return true;
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
