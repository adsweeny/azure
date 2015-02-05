using System;
using System.Data.Entity;

namespace adsweeny.Users
{
    public class Users
    {
        public int userID { get; set; }
        public string username { get; set; }
        public string first { get; set; }
        public string last { get; set; }
    }

    public class adsweeny : Users
    {
        public DbSet<adsweeny> Users { get; set; }
    }
}