using System;

namespace ExcelSqlDataExchange.Support
{
    public interface IRequireViewIdentification
    {
        Guid ViewID { get; }
    }
}
