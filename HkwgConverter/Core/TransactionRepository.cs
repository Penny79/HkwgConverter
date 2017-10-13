namespace HkwgConverter.Core
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;
    using Model;

    public partial class TransactionRepository : DbContext
    {
        public TransactionRepository()
            : base("name=TransactionRepository")
        {
        }

        public virtual DbSet<Transaction> Transactions { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
       
        }

        public int GetNextInputVersionNumer(DateTime deliveryday)
        {
            var result = this.Transactions.Where(x => x.DeliveryDate == deliveryday.Date);

            return result.Any() ? result.Max(x => x.Version) + 1 : 1;            
        }

        /// <summary>
        /// Returns the latest transaction in the system
        /// </summary>
        /// <param name="deliveryDay"></param>
        /// <returns>the latest transaction or null</returns>
        public Transaction GetLatest(DateTime deliveryDay)
        {
            var latestWorkFlow = this.Transactions.Where(x => x.DeliveryDate == deliveryDay)
                                                    .OrderByDescending(y => y.Version)
                                                        .FirstOrDefault();

            return latestWorkFlow;
        }
    }
}
