﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace StockTickerSparklines
{
    //  ------------------------------------------------------------------------
    //  This class was generated by a tool.
    //
    //  It was created by using the Paste Special tool on the Edit menu of the
    //  Visual Studio 2017 code editor to generate a class from a JSON string.
    //  Apart from adding this flower box, the only edit was the substitution of
    //  Rootobject with DailyTimeSeriesResponse, since we cannot afford to have
    //  all such objects in a namespace called Rootobject. (There is already one
    //  other object that was generated in this way in StockTickerSparklines.
    //  ------------------------------------------------------------------------

    public class DailyTimeSeriesResponse
    {
        public Meta_Data Meta_Data
        {
            get; set;
        }
        public Time_Series_Daily [ ] Time_Series_Daily
        {
            get; set;
        }
    }

    public class Meta_Data
    {
        public string Information
        {
            get; set;
        }
        public string Symbol
        {
            get; set;
        }
        public string LastRefreshed
        {
            get; set;
        }
        public string OutputSize
        {
            get; set;
        }
        public string TimeZone
        {
            get; set;
        }
    }

    public class Time_Series_Daily
    {
        public string Activity_Date
        {
            get; set;
        }
        public string Open
        {
            get; set;
        }
        public string High
        {
            get; set;
        }
        public string Low
        {
            get; set;
        }
        public string Close
        {
            get; set;
        }
        public string AdjustedClose
        {
            get; set;
        }
        public string Volume
        {
            get; set;
        }
        public string DividendAmount
        {
            get; set;
        }
        public string SplitCoefficient
        {
            get; set;
        }

        public double GetAdjustment ( )
        {
            double dblClose;
            double dblAdjustedClose;

            if ( double.TryParse ( Close , out dblClose ) )
            {
                if ( double.TryParse ( AdjustedClose , out dblAdjustedClose ) )
                {
                    return dblAdjustedClose - dblClose;
                }   // TRUE (anticipated outcome) block, if ( double.TryParse ( AdjustedClose , out dblAdjustedClose ) )
                else
                {
                    return 0;
                }   // FALSE (unanticipated outcome) block, if ( double.TryParse ( AdjustedClose , out dblAdjustedClose ) )
            }   // TRUE (anticipated outcome) block, if ( double.TryParse ( Close , out dblClose ) )
            else
            {
                return 0;
            }   // FALSE (unanticipated outcome) block, if ( double.TryParse ( Close , out dblClose ) )
        }   // public double GetAdjustment method

        internal static object ConvertToAppropriateType ( string pstrStringFromJSON )
        {
            DateTime dtmTemp;

            if ( DateTime.TryParse ( pstrStringFromJSON , out dtmTemp ) )
            {
                return dtmTemp;
            }   // TRUE (Input value is a DateTime.) block, if ( DateTime.TryParse ( pstrStringFromJSON , out dtmTemp ) )
            else
            {
                long lngTemp;

                if ( long.TryParse ( pstrStringFromJSON , out lngTemp ) )
                {
                    return lngTemp;
                }   // TRUE (Input value is a Long Integer.) block, if ( long.TryParse ( pstrStringFromJSON , out lngTemp ) )
                else
                {
                    double dblTemp;

                    if ( double.TryParse ( pstrStringFromJSON , out dblTemp ) )
                    {
                        return dblTemp;
                    }   // TRUE (Input value is a Double Precision floating point number.) block, if ( double.TryParse ( pstrStringFromJSON , out dblTemp ) )
                    else
                    {
                        return pstrStringFromJSON;
                    }   // FALSE (Input value is of another type) block, if ( double.TryParse ( pstrStringFromJSON , out dblTemp ) )
                }   // FALSE (Input value is of another type.) block, if ( long.TryParse ( pstrStringFromJSON , out lngTemp ) )
            }   // FALSE (Input value is of another type.) block, if ( DateTime.TryParse ( pstrStringFromJSON , out dtmTemp ) )
        }   // internal static object ConvertToAppropriateType
    }   // public class Time_Series_Daily
}   // partial namespace StockTickerSparklines