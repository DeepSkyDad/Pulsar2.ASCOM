﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using ASCOM.DeviceInterface;
using System.Collections;

namespace ASCOM.Pulsar2
{
    #region Rate class
    //
    // The Rate class implements IRate, and is used to hold values
    // for AxisRates. You do not need to change this class.
    //
    // The Guid attribute sets the CLSID for ASCOM.Pulsar2.Rate
    // The ClassInterface/None addribute prevents an empty interface called
    // _Rate from being created and used as the [default] interface
    //
    [Guid("37f6e6ff-6a50-4b74-8531-79884054ff14")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class Rate : ASCOM.DeviceInterface.IRate
    {
        private double maximum = 0;
        private double minimum = 0;

        //
        // Default constructor - Internal prevents public creation
        // of instances. These are values for AxisRates.
        //
        internal Rate(double minimum, double maximum)
        {
            this.maximum = maximum;
            this.minimum = minimum;
        }

        #region Implementation of IRate

        public void Dispose()
        {
            throw new System.NotImplementedException();
        }

        public double Maximum
        {
            get { return this.maximum; }
            set { this.maximum = value; }
        }

        public double Minimum
        {
            get { return this.minimum; }
            set { this.minimum = value; }
        }

        #endregion
    }
    #endregion

    #region AxisRates
    //
    // AxisRates is a strongly-typed collection that must be enumerable by
    // both COM and .NET. The IAxisRates and IEnumerable interfaces provide
    // this polymorphism. 
    //
    // The Guid attribute sets the CLSID for ASCOM.Pulsar2.AxisRates
    // The ClassInterface/None addribute prevents an empty interface called
    // _AxisRates from being created and used as the [default] interface
    //
    [Guid("aa7393c3-5ce1-4fb1-a95d-f53e4ac2f4bf")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class AxisRates : IAxisRates, IEnumerable
    {
        private TelescopeAxes axis;
        private readonly Rate[] rates;

        //
        // Constructor - Internal prevents public creation
        // of instances. Returned by Telescope.AxisRates.
        //
        internal AxisRates(TelescopeAxes axis)
        {
            this.axis = axis;
            //
            // This collection must hold zero or more Rate objects describing the 
            // rates of motion ranges for the Telescope.MoveAxis() method
            // that are supported by your driver. It is OK to leave this 
            // array empty, indicating that MoveAxis() is not supported.
            //
            // Note that we are constructing a rate array for the axis passed
            // to the constructor. Thus we switch() below, and each case should 
            // initialize the array for the rate for the selected axis.
            //
            switch (axis)
            {
                case TelescopeAxes.axisPrimary:
                    //RA rate in degress per second (minimum 0, maximum 4)
                    this.rates = new Rate[] { new Rate(0, 4) };
                    break;
                case TelescopeAxes.axisSecondary:
                    //DEC rate in degress per second (minimum 0, maximum 4)
                    this.rates = new Rate[] { new Rate(0, 4) };
                    break;
                case TelescopeAxes.axisTertiary:
                    // no rates - this is for field rotator
                    this.rates = new Rate[0];
                    break;
            }
        }

        #region IAxisRates Members

        public int Count
        {
            get { return this.rates.Length; }
        }

        public void Dispose()
        {
            throw new System.NotImplementedException();
        }

        public IEnumerator GetEnumerator()
        {
            return rates.GetEnumerator();
        }

        public IRate this[int index]
        {
            get { return this.rates[index - 1]; }	// 1-based
        }

        #endregion
    }
    #endregion

    #region TrackingRates
    //
    // TrackingRates is a strongly-typed collection that must be enumerable by
    // both COM and .NET. The ITrackingRates and IEnumerable interfaces provide
    // this polymorphism. 
    //
    // The Guid attribute sets the CLSID for ASCOM.Pulsar2.TrackingRates
    // The ClassInterface/None addribute prevents an empty interface called
    // _TrackingRates from being created and used as the [default] interface
    //
    [Guid("2753238d-dbc8-4a1d-9cb2-72f153192b5b")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class TrackingRates : ITrackingRates, IEnumerable, IEnumerator
    {
        private readonly DriveRates[] trackingRates;
        private static int pos = -1;

        //
        // Default constructor - Internal prevents public creation
        // of instances. Returned by Telescope.AxisRates.
        //
        internal TrackingRates()
        {
            //
            // This array must hold ONE or more DriveRates values, indicating
            // the tracking rates supported by your telescope. The one value
            // (tracking rate) that MUST be supported is driveSidereal!
            this.trackingRates = new[] {
                DriveRates.driveSidereal, 
                DriveRates.driveLunar,
                DriveRates.driveSolar
            };
        }

        #region ITrackingRates Members

        public int Count
        {
            get { return this.trackingRates.Length; }
        }

        public IEnumerator GetEnumerator()
        {
            pos = -1;
            return this as IEnumerator;
        }

        public void Dispose()
        {
            throw new System.NotImplementedException();
        }

        public DriveRates this[int index]
        {
            get { return this.trackingRates[index - 1]; }	// 1-based
        }

        #endregion

        #region IEnumerable members

        public object Current
        {
            get
            {
                if (pos < 0 || pos >= trackingRates.Length)
                {
                    throw new System.InvalidOperationException();
                }
                return trackingRates[pos];
            }
        }

        public bool MoveNext()
        {
            if (++pos >= trackingRates.Length)
            {
                return false;
            }
            return true;
        }

        public void Reset()
        {
            pos = -1;
        }
        #endregion
    }
    #endregion
}
