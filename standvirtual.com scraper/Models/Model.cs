namespace standvirtual.com_scraper.Models
{
    public class Model
    {
        public string Name { get; set; }
        public string SearchKey { get; set; }
        public string Id { get; set; }
        public override string ToString()
        {
            return Name;
        }

        public override bool Equals(object obj)
        {
            return SearchKey?.Equals(((Model)obj)?.SearchKey) ?? false;
        }
    }
}
