using System.Collections.Generic;

namespace standvirtual.com_scraper.Models
{
    public class Make
    {
        public string SearchKey { get; set; }
        public string Name { get; set; }
        public string Id { get; set; }
        public List<Model> Models { get; set; } = new List<Model>();
        public override string ToString()
        {
            return Name;
        }

        public override bool Equals(object obj)
        {
            if (obj.GetType() == typeof(Make))
                return SearchKey.Equals(((Make)obj).SearchKey);
            return base.Equals(obj);
        }
    }
}
