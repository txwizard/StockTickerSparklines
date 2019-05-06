using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockTickerSparklines
{
    /// <summary>
    /// The TickerSymbolMatches class holds the array of matches returned by a
    /// SYMBOL_SEARCH request.
    /// </summary>
    public class TickerSymbolMatches
    {
        public Bestmatch [ ] bestMatches
        {
            get; set;
        }
    }   // public class TickerSymbolMatches


    /// <summary>
    /// Each BestMatch item describes a stock ticker symbol that matches the
    /// SYMBOL_SEARCH request.
    /// </summary>
    public class Bestmatch
    {
        /// <summary>
        /// Ticker symbol
        /// </summary>
        public string _1symbol
        {   // The Visual Studio JSON parser generated this as _1symbol.
            get; set;
        }


        /// <summary>
        /// Issue name
        /// </summary>
        public string _2name
        {   // The Visual Studio JSON parser generated this as _2name.
            get; set;
        }


        /// <summary>
        /// Issue type
        /// </summary>
        public string _3type
        {   // The Visual Studio JSON parser generated this as _3type.
            get; set;
        }


        /// <summary>
        /// Region
        /// </summary>
        public string _4region
        {   // The Visual Studio JSON parser generated this as _4region.
            get; set;
        }


        /// <summary>
        /// Market opening time
        /// </summary>
        public string _5marketOpen
        {   // The Visual Studio JSON parser generated this as _5marketOpen.
            get; set;
        }


        /// <summary>
        /// Market closing time
        /// </summary>
        public string _6marketClose
        {   // The Visual Studio JSON parser generated this as _6marketClose.
            get; set;
        }

        /// <summary>
        /// Market time zone relative to which opening and closing times are
        /// specified
        /// </summary>
        public string _7timezone
        {   // The Visual Studio JSON parser generated this as _7timezone.
            get; set;
        }

        /// <summary>
        /// Currency in which price is given
        /// </summary>
        public string _8currency
        {   // The Visual Studio JSON parser generated this as _8currency.
            get; set;
        }

        /// <summary>
        /// Match score
        /// </summary>
        public string _9matchScore
        {   // The Visual Studio JSON parser generated this as _9matchScore.
            get; set;
        }
    }   // public class Bestmatch
}   // partial namespace StockTickerSparklines