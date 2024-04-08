using api_word2wp.Interfaces;
using api_word2wp.Models;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;
using RestSharp;
using Newtonsoft.Json;

namespace api_word2wp.Implements
{
    public class CategoryService : ICategoryService
    {
        public CategoryService() { }

        public async Task<List<WpCategory>> GetList()
        {
            try
            {
                var client = new RestClient("https://development.matbao.website/wp-json/mbwsapi/v1/category/list");
                var request = new RestRequest();
                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Api-Key", "MatBaoWS@1234");
                var response = await client.ExecutePostAsync(request);
                if (response.StatusCode != System.Net.HttpStatusCode.OK)
                {
                    return new List<WpCategory>();
                }
                Dictionary<string, WpCategory> categoryDict = JsonConvert.DeserializeObject<Dictionary<string, WpCategory>>(response.Content);
                List<WpCategory> categoryList = new List<WpCategory>(categoryDict.Values);
                return categoryList;
            }
            catch (Exception ex)
            {
                return new List<WpCategory>();
            }
        }
    }
}
