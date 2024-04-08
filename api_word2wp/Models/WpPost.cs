using System.Collections.Generic;

namespace api_word2wp.Models
{
    public class CreatePost
    {
        public List<string> Success { get; set; } = new List<string>();
        public List<string> Failed { get; set; } = new List<string>();
    }
}
