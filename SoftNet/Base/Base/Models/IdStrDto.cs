namespace Base.Models
{
    //for select option
    public class IdStrDto
    {
        public string Id { get; set; }
        public string Str { get; set; }
        public IdStrDto() { }
        public IdStrDto(string id, string str)
        {
            Id = id;
            Str = str;
        }
    }
}