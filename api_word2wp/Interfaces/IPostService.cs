using System.Threading.Tasks;

namespace api_word2wp.Interfaces
{
    public interface IPostService
    {
        Task<bool> AddPost(string content, string title, string thumbnail, string categories);
    }
}
