using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data.Models
{
    public class ShoppingCart
    {
        private readonly AnhCoderDbContext _context;
        public ShoppingCart(AnhCoderDbContext context)
        {
            _context = context;
        }
        public string ShoppingCartId { get; set; }
        public List<ShoppingCartItem> ShoppingCartItems { get; set; }
        
        public static ShoppingCart GetCart(IServiceProvider services)
        {
            ISession session = services.GetRequiredService<IHttpContextAccessor>()?.HttpContext.Session;
            var context = services.GetService<AnhCoderDbContext>();
            var cartId = session.GetString("CartId") ?? Guid.NewGuid().ToString();

            session.SetString("CartId", cartId);
            return new ShoppingCart(context) { ShoppingCartId = cartId };
        }
        public void AddToCart(ProductServices product,int amount)
        {
            var shoppingCartItem = _context.ShoppingCartItems.SingleOrDefault(s => s.ProductService.Id == product.Id && s.ShoppingCartId == ShoppingCartId);
            if (shoppingCartItem == null)
            {
                shoppingCartItem = new ShoppingCartItem
                {
                    ShoppingCartId = ShoppingCartId,
                    ProductService = product,
                    Amount = 1
                };
                _context.ShoppingCartItems.Add(shoppingCartItem);
            }
            else
            {
                shoppingCartItem.Amount++;
            }
            _context.SaveChanges(); 
        }
        public int RemoveFromCart(ProductServices product)
        {
            var shoppingCartIte = _context.ShoppingCartItems.SingleOrDefault(s => s.ProductService.Id == product.Id && s.ShoppingCartId == ShoppingCartId);
            var localAmount = 0;

            if (ShoppingCartItems != null)
            {
                if (shoppingCartIte.Amount > 1)
                {
                    shoppingCartIte.Amount--;
                    localAmount = shoppingCartIte.Amount;
                }
                else
                {
                    _context.ShoppingCartItems.Remove(shoppingCartIte);
                }
            }
            _context.SaveChanges();
            return localAmount;
        }
        public List<ShoppingCartItem> GetShoppingCartItems()
        {
            return ShoppingCartItems ??
                (ShoppingCartItems = _context.ShoppingCartItems.
                Where(c => c.ShoppingCartId == ShoppingCartId).
                Include(s => s.ProductService).
                ToList());
        }
        public void ClearCard()
        {
            var cardItems = _context.ShoppingCartItems.
                Where(card => card.ShoppingCartId == ShoppingCartId);
            _context.ShoppingCartItems.RemoveRange(cardItems);
            _context.SaveChanges();
        }
        public decimal GetShoppingCardTotal()
        {
            var total = _context.ShoppingCartItems.Where(c => c.ShoppingCartId == ShoppingCartId)
                .Select(c => c.ProductService.Price * c.Amount).Sum();
            return total;
        }
    }
}
