using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelAddIn1.Core
{
    public static class OfflineUpdateTrustCatalog
    {
        private const string ProductionPublicKey =
            "TTBMVN_OFFLINE_PUBLIC_KEY\n" +
            "schema=1\n" +
            "keyId=VFRCTVZOLU9GRkxJTkUtMjAyNi0wMQ==\n" +
            "algorithm=RSA-SHA256-PKCS1\n" +
            "modulus=oes43ED6TEEnJv5tmcF92Tlunx3fSkzUQFNc41uPGhrFRxsQL17Bq5o6kZXq0V7xj138xO7hJy2uvtUmMciDgSLzSfbt8eBre0u13UNmSlIXFACsn+zc4kw6IUrAGLwWHVBIFUcAqG3572mD80xGpMeDHz4k4Zn+1wvQGYO93ZuSTZ8FRtLuChSVWmxwRymiOhEbUWnmb/rLn4BixV7Uu+XwzGw7xTD6SPSGaBgN/OROvCvFHSLRaVxln03p5i2IzybP4DiwjOFnECQNVA4pFc+8Xx27ypo06LnxCrA0go6D1ifIZOU9sfZeE3r5tpH5R6HNkw13OvsWTAWE4H4dZ0GHDIYakvgJi4MWIHjqaRd/bp0qNfJs6sqC+GRhuR7g7QOvM2Sf7vydS250hTKmMxFlW9gR8tiiYMwRNEKkG1G2/1LhoQyXmNjyCurptIpkXOWVDOb/hpDf0ednc7xy6/DUsTcy+Ke8V5zPwxhGYGy5WxaBHrHPmV9JK7qfVanh\n" +
            "exponent=AQAB\n";

        private static readonly IReadOnlyList<OfflineUpdateTrustedKey> production =
            new ReadOnlyCollection<OfflineUpdateTrustedKey>(
                new[] { OfflineUpdatePublicKeySerializer.Deserialize(ProductionPublicKey) });

        public static IReadOnlyList<OfflineUpdateTrustedKey> Production => production;
    }
}
