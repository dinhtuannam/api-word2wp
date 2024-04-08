using api_word2wp.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace api_word2wp.Interfaces
{
    public interface ICategoryService
    {
        Task<List<WpCategory>> GetList();
    }
}
