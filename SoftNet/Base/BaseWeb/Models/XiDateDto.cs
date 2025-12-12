namespace BaseWeb.Models
{
    public class XiDateDto : XiBaseDto
    {
        public bool IsAutoDate { get; set; } = false;
        public int IsAutoDate_MathDay { get; set; } = 0;
    }
}