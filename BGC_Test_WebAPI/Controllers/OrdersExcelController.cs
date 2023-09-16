using BGC_Test_WebAPI.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;

namespace BGC_Test_WebAPI.Controllers
{
    [Route("api/ordersexcel")]
    [ApiController]
    public class OrdersExcelController : ControllerBase
    {
        private readonly string excelpath = "\\Excel\\Food sales.xlsx";
        private string excelFilePath = "";
        public OrdersExcelController(IHostingEnvironment hostingEnvironment)
        {
            excelFilePath = hostingEnvironment.ContentRootPath + excelpath;
        }
        // GET: api/orders
        [HttpGet]
        public ActionResult<IEnumerable<Order>> Get(
            DateTime? startDate, 
            DateTime? endDate, 
            DateTime? orderDate, 
            string? region, 
            string? city, 
            string? category, 
            string? product, 
            int? quantity, 
            decimal? unitPrice, 
            decimal? totalPrice)
        {
            // Read data from Excel file
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                var orders = new List<Order>();

                for (int row = 2; row <= rowCount; row++)
                {
                    orders.Add(new Order
                    {
                        Id = row,
                        OrderDate = DateTime.Parse(worksheet.Cells[row, 1].Value.ToString()),
                        Region = worksheet.Cells[row, 2].Value.ToString(),
                        City = worksheet.Cells[row, 3].Value.ToString(),
                        Category = worksheet.Cells[row, 4].Value.ToString(),
                        Product = worksheet.Cells[row, 5].Value.ToString(),
                        Quantity = int.Parse(worksheet.Cells[row, 6].Value.ToString()),
                        UnitPrice = decimal.Parse(worksheet.Cells[row, 7].Value.ToString()),
                        TotalPrice = decimal.Parse(worksheet.Cells[row, 8].Value.ToString())
                    });
                }

                // filter
                if (orderDate != null)
                    orders = orders.Where(t => t.OrderDate == orderDate).ToList();
                if (startDate != null) 
                    orders = orders.Where(t => t.OrderDate >= startDate).ToList();
                if (endDate != null) 
                    orders = orders.Where(t => t.OrderDate <= endDate).ToList();
                if (region != null) orders = orders.Where(t => t.Region.Contains(region)).ToList();
                if (city != null) orders = orders.Where(t => t.City.Contains(city)).ToList();
                if (category != null) orders = orders.Where(t => t.Category.Contains(category)).ToList();
                if (product != null) orders = orders.Where(t => t.Product.Contains(product)).ToList();
                if (quantity != null) orders = orders.Where(t => t.Quantity == quantity).ToList();
                if (unitPrice != null) orders = orders.Where(t => t.UnitPrice == unitPrice).ToList();
                if (totalPrice != null) orders = orders.Where(t => t.TotalPrice == totalPrice).ToList();

                return Ok(orders);
            }
        }

        // POST: api/orders
        [HttpPost]
        public IActionResult Post([FromBody] Order newOrder)
        {
            if (newOrder == null)
            {
                return BadRequest("Invalid data");
            }

            // Read existing data from Excel file
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                // Find the next available row for the new order
                int newRow = rowCount + 1;

                // Write the new order data to Excel
                worksheet.Cells[newRow, 1].Value = newOrder.OrderDate.ToString();
                worksheet.Cells[newRow, 2].Value = newOrder.Region;
                worksheet.Cells[newRow, 3].Value = newOrder.City;
                worksheet.Cells[newRow, 4].Value = newOrder.Category;
                worksheet.Cells[newRow, 5].Value = newOrder.Product;
                worksheet.Cells[newRow, 6].Value = newOrder.Quantity;
                worksheet.Cells[newRow, 7].Value = newOrder.UnitPrice;
                worksheet.Cells[newRow, 8].Value = newOrder.TotalPrice;

                // Save changes to Excel file
                package.Save();
            }

            return Ok(newOrder);
        }

        // PUT: api/orders/5
        [HttpPut("{id}")]
        public IActionResult Put(int id, [FromBody] Order updatedOrder)
        {
            if (updatedOrder == null)
            {
                return BadRequest("Invalid data");
            }

            // Read existing data from Excel file
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                // Check if the specified ID is valid
                if (id >= 2 && id <= rowCount)
                {
                    // Update the order data in Excel
                    worksheet.Cells[id, 1].Value = updatedOrder.OrderDate.ToString();
                    worksheet.Cells[id, 2].Value = updatedOrder.Region;
                    worksheet.Cells[id, 3].Value = updatedOrder.City;
                    worksheet.Cells[id, 4].Value = updatedOrder.Category;
                    worksheet.Cells[id, 5].Value = updatedOrder.Product;
                    worksheet.Cells[id, 6].Value = updatedOrder.Quantity;
                    worksheet.Cells[id, 7].Value = updatedOrder.UnitPrice;
                    worksheet.Cells[id, 8].Value = updatedOrder.TotalPrice;

                    // Save changes to Excel file
                    package.Save();

                    return Ok(updatedOrder);
                }
                else
                {
                    return NotFound("Order not found");
                }
            }
        }

        // DELETE: api/orders/5
        [HttpDelete("{id}")]
        public IActionResult Delete(int id)
        {
            // Read existing data from Excel file
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                // Check if the specified ID is valid
                if (id >= 2 && id <= rowCount)
                {
                    // Delete the row from Excel
                    worksheet.DeleteRow(id);

                    // Save changes to Excel file
                    package.Save();

                    return NoContent();
                }
                else
                {
                    return NotFound("Order not found");
                }
            }
        }
    }
}
