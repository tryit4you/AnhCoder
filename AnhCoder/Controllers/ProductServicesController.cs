using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using AnhCoder.Data;
using AnhCoder.Data.Models;

namespace AnhCoder.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProductServicesController : ControllerBase
    {
        private readonly AnhCoderDbContext _context;

        public ProductServicesController(AnhCoderDbContext context)
        {
            _context = context;
        }

        // GET: api/ProductServices
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ProductService>>> GetProductServices()
        {

            IEnumerable<ProductService> products =await _context.ProductServices.ToListAsync();
            return Ok(products);
        }

        // GET: api/ProductServices/5
        [HttpGet("{id}")]
        public async Task<ActionResult<ProductService>> GetProductService(string id)
        {
            var productService = await _context.ProductServices.FindAsync(id);

            if (productService == null)
            {
                return NotFound();
            }

            return productService;
        }

        // PUT: api/ProductServices/5
        [HttpPut("{id}")]
        public async Task<IActionResult> PutProductService(string id, ProductService productService)
        {
            if (id != productService.Id)
            {
                return BadRequest();
            }

            _context.Entry(productService).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ProductServiceExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return NoContent();
        }

        // POST: api/ProductServices
        [HttpPost]
        public async Task<ActionResult<ProductService>> PostProductService(ProductService productService)
        {
            _context.ProductServices.Add(productService);
            await _context.SaveChangesAsync();

            return CreatedAtAction("GetProductService", new { id = productService.Id }, productService);
        }

        // DELETE: api/ProductServices/5
        [HttpDelete("{id}")]
        public async Task<ActionResult<ProductService>> DeleteProductService(string id)
        {
            var productService = await _context.ProductServices.FindAsync(id);
            if (productService == null)
            {
                return NotFound();
            }

            _context.ProductServices.Remove(productService);
            await _context.SaveChangesAsync();

            return productService;
        }

        private bool ProductServiceExists(string id)
        {
            return _context.ProductServices.Any(e => e.Id == id);
        }
    }
}
