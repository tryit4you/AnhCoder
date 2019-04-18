using AnhCoder.Data.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data
{
    public class AnhCoderDbContext:DbContext
    {
        public AnhCoderDbContext(DbContextOptions options) : base(options)
        {

        }
        public virtual DbSet<ShoppingCartItem> ShoppingCartItems { get; set; }
        public virtual DbSet<Order> Orders { get; set; }
        public virtual DbSet<OrderDetail> OrderDetails { get; set; }
        public virtual DbSet<Contact> Contacts { get; set; }
        public virtual DbSet<Evaluation> Evaluations { get; set; }
        public virtual DbSet<FeedBack> FeedBacks { get; set; }
        public virtual DbSet<ProductService> ProductServices { get; set; }
        public virtual DbSet<Service> Services { get; set; }
        public virtual DbSet<Slide> Slides { get; set; }
        public virtual DbSet<Page> Pages { get; set; }
    }
}
