using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace StockTickerSparklines
{
    public class DailyTimeSeriesResponse
    {
        public MetaData Meta
        {
            get; set;
        }   // Read/Write Meta property


        public TimeSeriesDaily TimeSeries
        {
            get; set;
        }
    }   // public class TimeSeries property


    public class MetaData
    {
        public string Information
        {
            get; set;
        }   // Read/Write Information property


        public string Symbol
        {
            get; set;
        }   // Read/Write Symbol property


        public string LastRefreshed
        {
            get; set;
        }   // Read/Write LastRefreshed property


        public string OutputSize
        {
            get; set;
        }   // Read/Write OutputSize property


        public string TimeZone
        {
            get; set;
        }   // Read/Write TimeZone property
    }   // public class MetaData


    public class TimeSeriesDaily
    {
        public List<DailyResult> DailyResults
        {
            get; set;
        }   // Read/Write DailyResults property
    }   // public class TimeSeriesDaily

    public class DailyResult
    {
        public string Open
        {
            get; set;
        }   // Read/Write Open property


        public string High
        {
            get; set;
        }   // Read/Write High property


        public string Low
        {
            get; set;
        }   // Read/Write Low property


        public string Close
        {
            get; set;
        }   // Read/Write Close property


        public string AdjustedClose
        {
            get; set;
        }   // Read/Write AdjustedClose property


        public string Volume
        {
            get; set;
        }   // Read/Write Volume property


        public string DividendAmount
        {
            get; set;
        }   // Read/Write DividendAmount property


        public string SplitCoefficient
        {
            get; set;
        }   // Read/Write SplitCoefficient property
    }   // public class DailyResult
}   // partial namespace StockTickerSparklines