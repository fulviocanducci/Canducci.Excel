namespace Canducci.Console.Test
{
    public class People
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Salary { get; set; } = 10;
        public System.DateTime DateCreated { get; set; } = System.DateTime.Now;
    }
}
