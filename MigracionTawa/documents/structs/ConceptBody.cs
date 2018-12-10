using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents.structs {
    class ConceptBody {
        public readonly string parsedDate;
        private double total;

        public ConceptBody(string docDate, double initialSum) {
            parsedDate = docDate;
            total = initialSum;
        }

        public void incSum(double nextSum) {
            total += nextSum;
        }

        public double getTotalSum() {
            return total;
        }

        public DateTime getDate() {
            return DateTime.ParseExact(parsedDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
