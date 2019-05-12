using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockTickerSparklines
{
    class SparklineRowAssignments
    {
        private SparklineRowAssignments ( )
        {
        }

        public SparklineRowAssignments ( int pintOriginColumn , int pintTimeSeriesItemCount )
        {
            SparklineColumnIndex = pintOriginColumn + pintTimeSeriesItemCount + 1;
        }

        public int Open
        {
            get; set;
        }

        public int High
        {
            get; set;
        }

        public int Low
        {
            get; set;
        }

        public int Close
        {
            get; set;
        }

        public int AdjustedClose
        {
            get; set;
        }

        public int Adjustment
        {
            get; set;
        }

        public int Volume
        {
            get; set;
        }

        public int SparklineColumnIndex
        {
            get; private set;
        }

        public int [ ] GetArrayOfSparklineRows ( )
        {
            const int INDEX_OPEN = 0;
            const int INDEX_HIGH = 1;
            const int INDEX_LOW = 2;
            const int INDEX_CLOSE = 3;
            const int INDEX_ADJUSTEDCLOSE = 4;
            const int INDEX_ADJUSTMENT = 5;
            const int INDEX_VOLUME = 6;
            const int SPARKLINE_ROWS = 7;

            int [ ] raintSparklineRows = new int [ SPARKLINE_ROWS ];

            raintSparklineRows [ INDEX_OPEN ] = Open;
            raintSparklineRows [ INDEX_HIGH ] = High;
            raintSparklineRows [ INDEX_LOW ] = Low;
            raintSparklineRows [ INDEX_CLOSE ] = Close;
            raintSparklineRows [ INDEX_ADJUSTEDCLOSE ] = AdjustedClose;
            raintSparklineRows [ INDEX_ADJUSTMENT ] = Adjustment;
            raintSparklineRows [ INDEX_VOLUME ] = Volume;

            return raintSparklineRows;
        }   // public int [ ] GetArrayOfSparklineRows
    }   // class SparklineRowAssignments
}   // partial namespace StockTickerSparklines