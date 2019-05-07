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

        public override string ToString ( )
        {
            return string.Format (
                Properties.Resources.TPL_BESTMATCH_TOSTRING ,
                new string [ ]
                {
                    _1symbol ,          // Format Item 0: _1symbol (Symbol) = {0}
                    _2name ,            // Format Item 1: _2name (Name) = {1}
                    _3type ,            // Format Item 2: _3type (Type) = {2}
                    _4region ,          // Format Item 3: _4region (Region) = {3}
                    _5marketOpen ,      // Format Item 4: _5marketOpen (Market Opening Time) = {4}
                    _6marketClose ,     // Format Item 5: _6marketClose (Market Closing Time) = {5}
                    _7timezone ,        // Format Item 6: _7timezone (Time Zone) = {6}
                    _8currency ,        // Format Item 7: _8currency (Currency) = {7}
                    _9matchScore        // Format Item 8: __9matchScore (Match Score) = {8}
                } );
        }
    }   // public class Bestmatch
}   // partial namespace StockTickerSparklines