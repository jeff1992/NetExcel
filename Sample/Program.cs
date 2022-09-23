using NetExcel;

namespace Sample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Test1();
            Test2();
        }

        static void Test1()
        {
            Random random = new Random();
            Dictionary<string, IEnumerable<string>> dic = new Dictionary<string, IEnumerable<string>>();
            dic.Add("Fruit", new string[] { "Peach", "Plum", "Banana", "Pear" });
            dic.Add("Vegetable", new string[] { "Cabbage", "Potato", "Cucumber", "Bear" });

            //构造model
            var order = new
            {
                ProjectName = "Gray wolf's birthday party",
                Name = "Jeff",
                CreatedAt = DateTime.Now,
                BuyerName = "Bill",
                Cates = dic.Select(m => new
                {
                    Name = m.Key,
                    Items = m.Value.Select(n => new
                    {
                        Name = n,
                        Price = (decimal)random.Next(1, 100),
                        Amount = random.Next(1, 100)
                    }).ToList()
                })
            };
            ExcelTemplate render = new ExcelTemplate("./templates/purchase.xlsx");
            render.Values.Add("order", order);
            var fileName = DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            //bellow is the main method
            render.SaveAs(fileName, "123123");
        }

        static void Test2()
        {
            var model = new Model1
            {
                invoiceNo = "asdf",
                invoiceDate = DateTime.Now,
                deliveryNo = "1234",
                totalNumber = 334,
                totalAmount = 232,
                pos = new List<Model2>()
            };
            model.pos.Add(new Model2
            {
                deliveryFormNo = "1",
                totalAmount = 334.3M,
                details = new List<Model3>()
            });
            model.pos.Add(new Model2
            {
                deliveryFormNo = "2",
                totalAmount = 23232.1M,
                details = new List<Model3>()
            });
            model.pos[0].details.Add(new Model3
            {
                deliveryFormNo = "3"
            });
            model.pos[1].details.Add(new Model3
            {
                deliveryFormNo = "4"
            });

            var template = new ExcelTemplate("./templates/invoice.xlsx");
            template.Values.Add("model", model);
            var fileName = DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            template.SaveAs(fileName);
        }
    }


    public class Model1
    {
        public string invoiceNo { get; set; }
        public DateTime invoiceDate { get; set; }
        public string soldTo { get; set; }
        public string deliveryNo { get; set; }
        public decimal totalNumber { get; set; }
        public decimal totalAmount { get; set; }
        public List<Model2> pos { get; set; }

    }
    public class Model2
    {
        public string deliveryFormNo { get; set; }

        public decimal totalAmount { get; set; }
        public List<Model3> details { get; set; }
    }
    public class Model3
    {
        public string deliveryFormNo { get; set; }
        public string title { get; set; }
        public string mouldName { get; set; }
        public string partNo { get; set; }
        public string partName { get; set; }
        public int number { get; set; }
        public decimal price { get; set; }
        public decimal amount { get; set; }
    }
}