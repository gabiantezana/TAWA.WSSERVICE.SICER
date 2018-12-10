using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents.structs {
    class ConceptKey {
        public readonly int exCode;
        public readonly int etapa;
        public readonly string costingCode;
        public readonly string costingCode2;
        public readonly string costingCode3;
        public readonly string costingCode4;
        public readonly string costingCode5;
        public readonly string currency;
        public readonly string asunto;

        public ConceptKey(int xCode, int stage, string cc1, string cc2, string cc3, string cc4, string cc5, string curr, string memo) {
            exCode = xCode;
            etapa = stage;
            costingCode = cc1;
            costingCode2 = cc2;
            costingCode3 = cc3;
            costingCode4 = cc4;
            costingCode5 = cc5;
            currency = curr;
            asunto = memo;
        }

        public override bool Equals(object obj) {
            if (obj is ConceptKey) {
                ConceptKey parsedV = (ConceptKey)obj;
                return exCode.Equals(parsedV.exCode) && costingCode.Equals(parsedV.costingCode) && costingCode2.Equals(parsedV.costingCode2)
                     && costingCode3.Equals(parsedV.costingCode3) && costingCode4.Equals(parsedV.costingCode4)
                     && costingCode5.Equals(parsedV.costingCode5) && etapa.Equals(parsedV.etapa)
                     && asunto.Equals(parsedV.asunto);
            }
            return false;
        }

        public override int GetHashCode() {
            int hash = 13;
            hash += exCode.GetHashCode() * 7;
            hash += etapa.GetHashCode() * 7;
            hash += costingCode.GetHashCode() * 7;
            hash += costingCode2.GetHashCode() * 7;
            hash += costingCode3.GetHashCode() * 7;
            hash += costingCode4.GetHashCode() * 7;
            hash += costingCode5.GetHashCode() * 7;
            hash += asunto.GetHashCode() * 7;
            return hash;
        }
    }
}
