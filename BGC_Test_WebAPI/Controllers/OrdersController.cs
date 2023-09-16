using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using BGC_Test_WebAPI.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace BGC_Test_WebAPI.Controllers
{
    [Route("api/orders")]
    [ApiController]
    public class OrdersController : ControllerBase
    {
        private readonly BgcTestDbContext _dbContext;

        public OrdersController(BgcTestDbContext dbContext)
        {
            _dbContext = dbContext;
        }

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
            var orders = _dbContext.Orders.ToList();
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

        [HttpPost]
        public IActionResult Post([FromBody] Order newOrder)
        {
            if (newOrder == null)
            {
                return BadRequest("Invalid data");
            }

            _dbContext.Orders.Add(newOrder);
            _dbContext.SaveChanges();

            return Ok(newOrder);
        }

        [HttpPut("{id}")]
        public IActionResult Put(int id, [FromBody] Order updatedOrder)
        {
            if (updatedOrder == null)
            {
                return BadRequest("Invalid data");
            }

            var existingOrder = _dbContext.Orders.FirstOrDefault(o => o.Id == id);

            if (existingOrder == null)
            {
                return NotFound("Order not found");
            }

            existingOrder.OrderDate = updatedOrder.OrderDate;
            existingOrder.Region = updatedOrder.Region;
            existingOrder.City = updatedOrder.City;
            existingOrder.Category = updatedOrder.Category;
            existingOrder.Product = updatedOrder.Product;
            existingOrder.Quantity = updatedOrder.Quantity;
            existingOrder.UnitPrice = updatedOrder.UnitPrice;
            existingOrder.TotalPrice = updatedOrder.TotalPrice;

            _dbContext.SaveChanges();

            return Ok(existingOrder);
        }

        [HttpDelete("{id}")]
        public IActionResult Delete(int id)
        {
            var orderToDelete = _dbContext.Orders.FirstOrDefault(o => o.Id == id);

            if (orderToDelete == null)
            {
                return NotFound("Order not found");
            }

            _dbContext.Orders.Remove(orderToDelete);
            _dbContext.SaveChanges();

            return NoContent();
        }
    }
}
