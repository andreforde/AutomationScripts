using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class dal
    {
        public static void RemovePreviousOccurences()
        {
            using (var db = new XEEEntities())
            {
                List<Table> ta = db.Tables.ToList();
                foreach (Table i in ta)
                {
                    db.Tables.Remove(i);
                }
                db.SaveChanges();
            }
        }

        public static void SaveTableData(Table table)
        {
            using (var db = new XEEEntities())
            {
                db.Tables.Add(table);
                db.SaveChanges();
            }
        }

        public static List<Table> FetchTableDataWhereLocIsNotNull()
        {
            using (var db = new XEEEntities())
            {
                return db.Tables.Where(s => s.location != null).ToList();
            }
        }

        public static List<Table> FetchTableDataWhereLocIsNull()
        {
            using (var db = new XEEEntities())
            {
                return db.Tables.Where(s => s.location == null).ToList();
            }
        }
    }
}
