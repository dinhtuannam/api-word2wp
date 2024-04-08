namespace api_word2wp.Models
{
    public class WpCategory
    {
        public int term_id { get; set; }
        public string name { get; set; } = "";
        public string slug { get; set; } = "";
        public int term_group { get; set; }
        public int term_taxonomy_id { get; set; }
        public string taxonomy { get; set; } = "";
        public string description { get; set; } = "";
        public int parent { get; set; }
        public int count { get; set; }
        public string filter { get; set; } = "";
        public int cat_ID { get; set; }
        public int category_count { get; set; }
        public string category_description { get; set; } = "";
        public string cat_name { get; set; } = "";
        public string category_nicename { get; set; } = "";
        public int category_parent { get; set; }
    }
}
