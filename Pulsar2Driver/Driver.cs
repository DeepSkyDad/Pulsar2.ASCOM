//tabs=4
// --------------------------------------------------------------------------------
//
// ASCOM Telescope driver for Pulsar2
//
// Description:	ASCOM driver for Gemini Pulsar2
//
// Implements:	ASCOM Telescope interface version: <To be completed by driver developer>
// Author:		Pavle Gartner <pavle.gartner@gmail.com>
//
// Edit Log:
//
// Date			Who	            Vers	Description
// -----------	---	            -----	-------------------------------------------------------
// 01-12-2015	Pavle Gartner	6.0.0	Initial edit, created from ASCOM driver template
// --------------------------------------------------------------------------------------------
//


// This is used to define code in the template that is specific to one class implementation
// unused code canbe deleted and this definition removed.
#define Telescope

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.Utilities;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Collections;
using System.Threading;

namespace ASCOM.Pulsar2
{
    //
    // Your driver's DeviceID is ASCOM.Pulsar2.Telescope
    //
    // The Guid attribute sets the CLSID for ASCOM.Pulsar2.Telescope
    // The ClassInterface/None addribute prevents an empty interface called
    // _Pulsar2 from being created and used as the [default] interface
    //
    // TODO Replace the not implemented exceptions with code to implement the function or
    // throw the appropriate ASCOM exception.
    //

    /// <summary>
    /// ASCOM Telescope Driver for Pulsar2.
    /// </summary>
    [Guid("3dbb211f-df0e-4d6a-8cb7-efcd112630e0")]
    [ClassInterface(ClassInterfaceType.None)]
    public partial class Telescope : ITelescopeV3
    {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.Pulsar2.Telescope";
        // TODO Change the descriptive string for your driver then remove this line
        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "ASCOM Telescope Driver for Pulsar2.";

        internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal static string comPortDefault = "COM1";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";
        internal static string legacyPulseGuideProfileName = "Legacy Pulse Guide";
        internal static string legacyPulseGuideDefault = "false";
        internal static string commandTimeoutProfileName = "Command Timeout";
        internal static int commandTimeoutDefault = 5000;

        internal static string comPort; // Variables to hold the currrent device configuration
        internal static bool traceState;
        internal static bool legacyPulseGuide;
        internal static int commandTimeout;

        //buffered values
        internal static readonly object buffLock = new object();
        internal static DriveRates? bufferDriveRate = null;
        internal static bool buffIsPulseGuiding = false;

        //serial communication
        private Serial serial;

        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Private variable to hold the trace logger object (creates a diagnostic log file with information that you specify)
        /// </summary>
        private TraceLogger tl;
        private TraceLogger tls;

        /// <summary>
        /// Initializes a new instance of the <see cref="Pulsar2"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public Telescope()
        {
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl = new TraceLogger("", "Pulsar2");
            tl.Enabled = traceState;
            tl.LogMessage("Telescope", "Starting initialisation");

            tls = new TraceLogger("", "Pulsar2.Serial");
            tls.Enabled = traceState;

            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object

            tl.LogMessage("Telescope", "Completed initialisation");
        }


        //
        // PUBLIC COM INTERFACE ITelescopeV3 IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog()
        {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
                System.Windows.Forms.MessageBox.Show("Already connected, just press OK");

            using (SetupDialogForm F = new SetupDialogForm())
            {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        public ArrayList SupportedActions
        {
            get
            {
                //TODO
                tl.LogMessage("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            //TODO
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
            throw new MethodNotImplementedException();
        }

        public void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");
            this.CommandString(command, raw, true);
        }

        public bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");
            string ret = CommandString(command, raw, false);
            if (ret == "Ok" || ret == "1")
                return true;
            return false;
        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            return CommandString(command, raw, false);
        }

        public void Dispose()
        {
            // Clean up the tracelogger and util objects
            Connected = false;
            tl.Enabled = false;
            tl.Dispose();
            tl = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;
        }

        public bool Connected
        {
            get
            {
                tl.LogMessage("Connected Get", IsConnected.ToString());
                return IsConnected;
            }
            set
            {
                tl.LogMessage("Connected Set", value.ToString());
                if (value == IsConnected)
                    return;

                if (value)
                {
                    connectedState = true;
                    tl.LogMessage("Connected Set", "Connecting to port " + comPort);
                    serial = new Serial();
                    serial.Speed = SerialSpeed.ps9600;
                    serial.PortName = comPort;
                    serial.Connected = true;
                    serial.ReceiveTimeoutMs = commandTimeout;
                    CheckVersion();
                }
                else
                {
                    connectedState = false;
                    tl.LogMessage("Connected Set", "Disconnecting from port " + comPort);
                    serial.Connected = false;
                    serial.Dispose();
                }
            }
        }

        public string Description
        {
            get
            {
                tl.LogMessage("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverInfo = "Pulsar2 ASCOM driver by Pavle Gartner. Version: " + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                tl.LogMessage("InterfaceVersion Get", "3");
                return Convert.ToInt16("3");
            }
        }

        public string Name
        {
            get
            {
                string name = "Pulsar2 ASCOM";
                tl.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region ITelescope Implementation
        public void AbortSlew()
        {
            CheckConnected("AbortSlew");
            if (AtPark)
                throw new DriverException("Cannot abort slew - telescope is parked!");
            if (!Slewing)
            {
                tl.LogMessage("AbortSlew", "Not slewing!");
                return;
            }
                
            CommandBlind("Q");
            tl.LogMessage("AbortSlew", "Slewing aborted");
        }

        public AlignmentModes AlignmentMode
        {
            get
            {
                CheckConnected("AlignmentMode");
                AlignmentModes alignmentMode = AlignmentModes.algGermanPolar;
                var m = CommandString("YGM");
                if(m == "1") {
                    alignmentMode = AlignmentModes.algGermanPolar;
                } else if(m == "2") {
                    alignmentMode = AlignmentModes.algPolar;
                } else if(m == "3") {
                    alignmentMode = AlignmentModes.algAltAz;
                } else {
                    new DriverException("Unknown mount type" + m);
                }
                
                tl.LogMessage("AlignmentMode", alignmentMode.ToString());
                return alignmentMode;
            }
        }

        public double Altitude
        {
            get
            {
                CheckConnected("Altitude");
                var r = utilities.DMSToDegrees(CommandString("GA"));
                tl.LogMessage("Altitude Get", r.ToString());
                return r;
            }
        }

        public double ApertureArea
        {
            get
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("ApertureArea Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureArea", false);
            }
        }

        public double ApertureDiameter
        {
            get
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("ApertureDiameter Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureDiameter", false);
            }
        }

        public bool AtHome
        {
            get
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("AtHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool AtPark
        {
            get
            {
                var r = CommandBool("YGk");
                tl.LogMessage("AtPark", "Get - " + r.ToString());
                return r;
            }
        }

        public IAxisRates AxisRates(TelescopeAxes Axis)
        {
            tl.LogMessage("AxisRates", "Get - " + Axis.ToString());
            return new AxisRates(Axis);
        }

        public double Azimuth
        {
            get
            {
                CheckConnected("Azimuth");
                var r = utilities.DMSToDegrees(CommandString("GZ"));
                tl.LogMessage("Azimuth Get", r.ToString());
                return r;
            }
        }

        public bool CanFindHome
        {
            get
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("CanFindHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanMoveAxis(TelescopeAxes Axis)
        {
            tl.LogMessage("CanMoveAxis", "Get - " + Axis.ToString());
            switch (Axis)
            {
                case TelescopeAxes.axisPrimary: return true;
                case TelescopeAxes.axisSecondary: return true;
                case TelescopeAxes.axisTertiary: return false;
                default: throw new InvalidValueException("CanMoveAxis", Axis.ToString(), "0 to 2");
            }
        }

        public bool CanPark
        {
            get
            {
                tl.LogMessage("CanPark", "Get - " + true.ToString());
                return true;            
            }
        }

        public bool CanPulseGuide
        {
            get
            {
                tl.LogMessage("CanPulseGuide", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSetDeclinationRate
        {
            get
            {
                tl.LogMessage("CanSetDeclinationRate", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSetGuideRates
        {
            get
            {
                //TODO - on hold for now, check commented method DriverExtension.SetPulseGuideRate for details
                tl.LogMessage("CanSetGuideRates", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetPark
        {
            get
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("CanSetPark", "Get - " + true.ToString());
                return false;
            }
        }

        public bool CanSetPierSide
        {
            get
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("CanSetPierSide", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetRightAscensionRate
        {
            get
            {
                tl.LogMessage("CanSetRightAscensionRate", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSetTracking
        {
            get
            {
                tl.LogMessage("CanSetTracking", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSlew
        {
            get
            {
                tl.LogMessage("CanSlew", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSlewAltAz
        {
            get
            {
                //TODO - on hold for now
                tl.LogMessage("CanSlewAltAz", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAltAzAsync
        {
            get
            {
                //TODO - on hold for now
                tl.LogMessage("CanSlewAltAzAsync", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAsync
        {
            get
            {
                tl.LogMessage("CanSlewAsync", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSync
        {
            get
            {
                tl.LogMessage("CanSync", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSyncAltAz
        {
            get
            {
                //TODO - on hold for now
                tl.LogMessage("CanSyncAltAz", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanUnpark
        {
            get
            {
                tl.LogMessage("CanUnpark", "Get - " + true.ToString());
                return true;
            }
        }

        public double Declination
        {
            get
            {
                CheckConnected("Declination Get");
                var r = utilities.DMSToDegrees(CommandString("GD"));
                tl.LogMessage("Declination Get", r.ToString());
                return r;
            }
        }

        public double DeclinationRate
        {
            get
            {
                CheckConnected("DeclinationRate Get");
                var r = Helper.RadPerMin2ArcsecPerSec(ParseYGSDouble(CommandString("YGZ"))[1]);
                tl.LogMessage("DeclinationRate", "Get - " + r.ToString());
                return r;
            }
            set
            {
                CheckConnected("DeclinationRate Set");
                tl.LogMessage("DeclinationRate Set", value.ToString());
                var dec = Helper.ArcsecPerSec2RadPerMin(value);
                if(Math.Abs(dec) > 4.1887902)
                    throw new DriverException(string.Format("Invalid declination rate - {0} rad/min", dec.ToString()));
                var currentRa = ParseYGS(CommandString("YGZ"))[0];
                CommandBlind(string.Format("YSZ{0},{1}", currentRa, FormatYSZRadMinDouble(dec)));
            }
        }

        public PierSide DestinationSideOfPier(double RightAscension, double Declination)
        {
            //Unimplemented for now (confirmed by Andras)
            tl.LogMessage("DestinationSideOfPier", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("DestinationSideOfPier");
        }

        public bool DoesRefraction
        {
            get
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("DoesRefraction Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", false);
            }
            set
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("DoesRefraction Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", true);
            }
        }

        public EquatorialCoordinateType EquatorialSystem
        {
            get
            {
                EquatorialCoordinateType equatorialSystem = EquatorialCoordinateType.equLocalTopocentric;
                tl.LogMessage("DeclinationRate", "Get - " + equatorialSystem.ToString());
                return equatorialSystem;
            }
        }

        public void FindHome()
        {
            //Unimplemented (confirmed by Andras)
            tl.LogMessage("FindHome", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("FindHome");
        }

        public double FocalLength
        {
            get
            {
                tl.LogMessage("FocalLength Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("FocalLength", false);
            }
        }

        public double GuideRateDeclination
        {
            get
            {
                //TODO - on hold for now, check commented method DriverExtension.SetPulseGuideRate for details
                tl.LogMessage("GuideRateDeclination Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateDeclination", false);
            }
            set
            {

                //TODO - on hold for now, check commented method DriverExtension.SetPulseGuideRate for details
                tl.LogMessage("GuideRateDeclination Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateDeclination", true);
            }
        }

        public double GuideRateRightAscension
        {
            get
            {
                //TODO - on hold for now, check commented method DriverExtension.SetPulseGuideRate for details
                tl.LogMessage("GuideRateRightAscension Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateRightAscension", false);
            }
            set
            {
                //TODO - on hold for now, check commented method DriverExtension.SetPulseGuideRate for details
                tl.LogMessage("GuideRateRightAscension Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateRightAscension", true);
            }
        }

        public bool IsPulseGuiding
        {
            get
            {
                //return false - not timeout needed during pulse
                tl.LogMessage("IsPulseGuiding Get", buffIsPulseGuiding.ToString());
                return buffIsPulseGuiding;
            }
        }

        public void MoveAxis(TelescopeAxes Axis, double Rate)
        {
            CheckConnected("MoveAxis");
            if(!CanMoveAxis(Axis))
                throw new InvalidValueException("Axis " + Axis.ToString() + " not supported");
            
            if (AtPark)
                throw new InvalidValueException("Cannot move axis - telescope parked");

            //read current rates to preserve unchanged axis (LX200 commands work with both axes at once)
            var rateRange = AxisRates(Axis)[1];
            if(Math.Abs(Rate) < rateRange.Minimum || Math.Abs(Rate) > rateRange.Maximum)
                throw new InvalidValueException(string.Format("Invalid {0} rate - {1}", Axis.ToString(), Rate));

            tl.LogMessage("MoveAxis", string.Format("axis: {0}, rate: {1}", Axis.ToString(), Rate.ToString()));
            var currentRates = ParseYGS(CommandString("YGZ"));
            switch (Axis)
            {
                case TelescopeAxes.axisPrimary:
                    //positive is due West, negative due East
                    CommandBlind(string.Format("YSZ{0},{1}", FormatYSZRadMinDouble(Helper.DegPerSec2RadPerMin(Rate)), currentRates[1]));
                    break;
                case TelescopeAxes.axisSecondary:
                    //positive is due North, negative is due South
                    CommandBlind(string.Format("YSZ{0},{1}", currentRates[0], FormatYSZRadMinDouble(Helper.DegPerSec2RadPerMin(Rate))));
                    break;
            }
        }

        public void Park()
        {
            CheckConnected("Park");
            if (AtPark || IsMountParking)
                return;
            if (!CommandBool("YH"))
                throw new DriverException("Park failed!");
        }

        public void PulseGuide(GuideDirections Direction, int Duration)
        {
            CheckConnected("PulseGuide");

            if(AtPark)
                throw new DriverException("Cannot pulse guide - telescope is parked!");

            //asked by Andras - ignore less than bottom
            /*if (Duration < 10)
            {
                tl.LogMessage("PulseGuide IGNORED (less than 10ms)", string.Format("Direction: {0}, Rate: {1}", Direction.ToString(), Duration.ToString()));
                return;
            }*/
                
            //if (Duration > 9990)
            //    throw new DriverException(string.Format("Cannot pulse guide - invalid duration of {0} (min: 10ms, max:9990ms)", Duration));

            tl.LogMessage("PulseGuide", string.Format("Direction: {0}, Rate: {1}", Direction.ToString(), Duration.ToString()));
            var direction = string.Empty;
            switch (Direction)
            {
                case GuideDirections.guideNorth:
                    direction = "n";
                    break;
                case GuideDirections.guideSouth:
                    direction = "s";
                    break;
                case GuideDirections.guideEast:
                    direction = "e";
                    break;
                case GuideDirections.guideWest:
                    direction = "w";
                    break;
            }

            if (legacyPulseGuide)
            {
                //legacy PulseGuide, needed for MaximDL calibration (it does not work with Mg command)
                buffIsPulseGuiding = true;
                //hack requested by Andras -> set rate to guide explicitly due to firmware bug (should be set automatically when Mg is called)
                CommandBlind("RG");
                CommandBlind("M" + direction);
                utilities.WaitForMilliseconds(Duration);
                CommandBlind("Q");
                buffIsPulseGuiding = false;
            }
            else
            {
                //new command for guiding - does not work with MaximDL calibration
                buffIsPulseGuiding = false;
                //pulsar duration is in 10s of seconds (so e.g. for 200ms duration send 20 -> "Mgw203")
                CommandBlind(string.Format("Mg{0}{1}3", direction, (Duration / 10).ToString("000")));
            }
            
        }

        public double RightAscension
        {
            get
            {
                CheckConnected("RightAscension");
                var r = utilities.HMSToHours(CommandString("GR"));
                tl.LogMessage("RightAscension Get", r.ToString());
                return r;
            }
        }

        public double RightAscensionRate
        {
            get
            {
                CheckConnected("RightAscensionRate Get");
                var r = Helper.RadPerMin2SecPerSiderealSec(ParseYGSDouble(CommandString("YGZ"))[0]);
                tl.LogMessage("RightAscensionRate", "Get - " + r.ToString());
                return r;
            }
            set
            {
                CheckConnected("RightAscensionRate Set");
                tl.LogMessage("RightAscensionRate Set", value.ToString());
                var ra = Helper.SecPerSiderealSec2RadPerMin(value);
                if (Math.Abs(ra) > 4.1887902)
                    throw new DriverException(string.Format("Invalid ascension rate - {0} rad/min", ra.ToString()));
                var currentDec = ParseYGS(CommandString("YGZ"))[1];
                CommandBlind(string.Format("YSZ{0},{1}", FormatYSZRadMinDouble(ra), currentDec));
            }
        }

        public void SetPark()
        {
            //Unimplemented (confirmed by Andras)
            tl.LogMessage("SetPark", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SetPark");
        }

        public PierSide SideOfPier
        {
            get
            {
                CheckConnected("SideOfPier");

                //requested by Andras, 22.12.2015 - if fork, always return East
                if (AlignmentMode == AlignmentModes.algPolar)
                    return PierSide.pierEast;

                //https://en.wikipedia.org/wiki/Hour_angle
                //basically we rely on the HA value, BUT GW from Pulsar2 is an exception -> tells you if pointing state is normal (0) or not (1 -> flip is needed)
                var gw = CommandString("YGW");
                var ha = SiderealTime - RightAscension;

                /*
                 revised version by Andras, 3.12.2017
                    if GW=0 AND

                    0>=HA>-12 (or 24>HA>12 in the 24h version) then side of pier normal(west)

                    if GW=1 AND

                    0>=HA>-12 (or 24>HA>12 in the 24h version) then side of pier flipped (east)

                    if GW=0 AND

                    0<HA<=12 then side of pier flipped (east)

                    if GW=1 AND

                    0<HA<=12 then side of pier normal (west)"
                 */

                PierSide ps = PierSide.pierUnknown;
                var HA = SiderealTime - RightAscension;
                if(HA > 12)
                {
                    HA -= 24;
                } else if(HA < -11.9999999)
                {
                    HA += 24;
                }

                if(gw == "0" && 0>=HA && HA>-12)
                {
                    ps = PierSide.pierWest; //normal
                } else if(gw == "1" && 0 >= HA && HA > -12)
                {
                    ps = PierSide.pierEast; //flipped
                } else if(gw == "0" && 0<HA && HA<=12)
                {
                    ps = PierSide.pierEast; //flipped
                }
                else if (gw == "1" && 0 < HA && HA <= 12)
                {
                    ps = PierSide.pierWest; //normal
                }

                tl.LogMessage("SideOfPier", "Get - " + ps.ToString());
                return ps;
            }
            set
            {
                tl.LogMessage("SideOfPier Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SideOfPier", true);
            }
        }

        public double SiderealTime
        {
            get
            {
                CheckConnected("SiderealTime");
                double siderealTime = utilities.HMSToHours(this.CommandString("GS"));
                tl.LogMessage("SiderealTime", "Get - " + siderealTime.ToString());
                return siderealTime;
            }
        }

        public double SiteElevation
        {
            get
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("SiteElevation Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteElevation", false);
            }
            set
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("SiteElevation Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteElevation", true);
            }
        }

        public double SiteLatitude
        {
            get
            {
                CheckConnected("SiteLatitude");
                var r = utilities.DMSToDegrees(CommandString("Gt"));
                tl.LogMessage("SiteLatitude Get", r.ToString());
                return r;
            }
            set
            {
                if ((value < -90.0) || (value > 90.0))
                {
                    throw new InvalidValueException("Invalid TargetDeclination value " + value.ToString());
                }
                CheckConnected("SiteLatitude");
                tl.LogMessage("SiteLatitude Set", value.ToString());
                CommandBlind(string.Format("St {0}", utilities.DegreesToDM(value, "*", string.Empty)));
            }
        }

        public double SiteLongitude
        {
            get
            {
                CheckConnected("SiteLongitude");
                var gg = utilities.DMSToDegrees(CommandString("Gg"));
                if (gg > 180)
                    gg -= 360; //substract 360 if over 180
                gg *= -1; //negate final value (Pulsar2 has negative to east)
                tl.LogMessage("SiteLongitude Get", gg.ToString());
                return gg;
            }
            set
            {
                if ((value < -180) || (value > 180.0))
                {
                    throw new InvalidValueException("Invalid TargetDeclination value " + value.ToString());
                }
                CheckConnected("SiteLongitude");
                tl.LogMessage("SiteLongitude Set", value.ToString());

                var gg = value;
                gg *= -1;
                if (gg < 0)
                    gg += 360;

                CommandBlind(string.Format("Sg {0}", utilities.DegreesToDM(gg, "*", string.Empty)));
            }
        }

        public short SlewSettleTime
        {
            get
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("SlewSettleTime Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", false);
            }
            set
            {
                //Unimplemented (confirmed by Andras)
                tl.LogMessage("SlewSettleTime Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", true);
            }
        }

        public void SlewToAltAz(double Azimuth, double Altitude)
        {
            //TODO - on hold for now
            tl.LogMessage("SlewToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAz");
        }

        public void SlewToAltAzAsync(double Azimuth, double Altitude)
        {
            //TODO - on hold for now
            tl.LogMessage("SlewToAltAzAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAzAsync");
        }

        public void SlewToCoordinates(double RightAscension, double Declination)
        {
            CheckConnected("SlewToCoordinates");
            tl.LogMessage("SlewToCoordinates", string.Format("RA: {0}, DEC: {1}", RightAscension.ToString(), Declination.ToString()));
            TargetRightAscension = RightAscension;
            TargetDeclination = Declination;
            SlewToTarget();
        }

        public void SlewToCoordinatesAsync(double RightAscension, double Declination)
        {
            CheckConnected("SlewToCoordinatesAsync");
            tl.LogMessage("SlewToCoordinatesAsync", string.Format("RA: {0}, DEC: {1}", RightAscension.ToString(), Declination.ToString()));
            TargetRightAscension = RightAscension;
            TargetDeclination = Declination;
            SlewToTargetAsync();
        }

        public void SlewToTarget()
        {
            CheckConnected("SlewToTarget");
            tl.LogMessage("SlewToTarget", string.Empty);
            SlewToTargetAsync();
            while (Slewing)
            {
                utilities.WaitForMilliseconds(500);
            }
        }

        public void SlewToTargetAsync()
        {
            CheckConnected("SlewToTargetAsync");
            if (AtPark)
                throw new DriverException("SlewToTargetAsync - telescope is parked");
            if(!Tracking)
                throw new DriverException("SlewToTargetAsync - telescope is not tracking");

            tl.LogMessage("SlewToTargetAsync", string.Empty);
            if(CommandString("MS") == "1")
                throw new DriverException("SlewToTargetAsync - slewing failed");
        }

        public bool Slewing
        {
            get
            {
                var r = CommandBool("YGi");
                tl.LogMessage("Slewing Get", r.ToString());
                return r;
            }
        }

        public void SyncToAltAz(double Azimuth, double Altitude)
        {
            //TODO - on hold for now
            tl.LogMessage("SyncToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToAltAz");
        }

        public void SyncToCoordinates(double RightAscension, double Declination)
        {
            CheckConnected("SyncToCoordinates");
            TargetRightAscension = RightAscension;
            TargetDeclination = Declination;
            SyncToTarget();
        }

        public void SyncToTarget()
        {
            CheckConnected("SyncToTarget");
            if (AtPark)
                throw new DriverException("SyncToTarget - telescope is parked");
            if (!Tracking)
                throw new DriverException("SyncToTarget - telescope is not tracking");
            utilities.WaitForMilliseconds(300);
            CommandString("CM");
        }

        public double TargetDeclination
        {
            get
            {
                CheckConnected("TargetDeclination Get");
                var r = utilities.DMSToDegrees(CommandString("Gd"));
                tl.LogMessage("TargetDeclination Get", r.ToString());
                return r;
            }
            set
            {
                CheckConnected("TargetDeclination Set");
                if ((value < -90.0) | (value > 90.0))
                {
                    throw new InvalidValueException("Invalid TargetDeclination value " + value.ToString());
                }

                tl.LogMessage("TargetDeclination Set", value.ToString());
                var cmd = utilities.DegreesToDMS(value, "*", ":", string.Empty);
                //if (cmd[0] != '-')
                //    cmd = "+" + cmd;
                cmd = string.Format("Sd {0}", cmd);
                if (!CommandBool(cmd))
                {
                    tl.LogMessage("TargetDeclination Set", "Failed - retry");
                    //requested by Andras Dan - retry if Sd fails
                    if (!CommandBool(cmd))
                    {
                        tl.LogMessage("TargetDeclination Set", "Failed twice - give up");
                        throw new DriverException(string.Format("Set TargetDeclination command failed ({0})", cmd));
                    }
                }
                    
            }
        }

        public double TargetRightAscension
        {
            get
            {
                CheckConnected("TargetRightAscension");
                var r = utilities.DMSToDegrees(CommandString("Gr"));
                tl.LogMessage("TargetRightAscension Get", r.ToString());
                return r;
            }
            set
            {
                CheckConnected("TargetRightAscension Set");
                if ((value < 0.0) | (value >= 24.0))
                {
                    throw new InvalidValueException("Invalid TargetRightAscension value " + value.ToString());
                }
                tl.LogMessage("TargetRightAscension Set", value.ToString());
                var cmd = string.Format("Sr {0}", utilities.HoursToHMS(value, ":", ":"));
                if (!CommandBool(cmd))
                {
                    tl.LogMessage("TargetRightAscension Set", "Failed - retry");
                    //requested by Andras Dan - retry if Sr fails
                    if (!CommandBool(cmd))
                    {
                        tl.LogMessage("TargetRightAscension Set", "Failed twice - give up");
                        throw new DriverException(string.Format("Set TargetRightAscension command failed ({0})", cmd));
                    }
                }
            }
        }

        public bool Tracking
        {
            get
            {
                CheckConnected("Tracking Get");
                var tracking = CommandString("YGS").Split(',')[0] != "0";
                tl.LogMessage("Tracking", "Get - " + tracking.ToString());
                return tracking;
            }
            set
            {
                CheckConnected("Tracking Set");
                tl.LogMessage("Tracking Set", value.ToString());
                
                if (value)
                {
                    //apply rate
                    CommandBlind(string.Format("YSS{0},0", Helper.GetPulsarRateFromASCOMRate(TrackingRate)));
                }
                else
                {
                    //stop tracking (sidereal)
                    CommandBlind("YSS0,0");
                }
            }
       } 

        public DriveRates TrackingRate
        {
            get
            {
                CheckConnected("TrackingRate Get");
                if (bufferDriveRate == null) {
                    //initial read
                    var rate = CommandString("YGS").Split(',')[0];
                    lock (buffLock)
                    {
                        bufferDriveRate = rate == "0" ? DriveRates.driveSidereal : Helper.GetASCOMRateFromPulsarRate(Convert.ToInt32(rate));
                    }
                }
                    
                tl.LogMessage("TrackingRate Get", ((int)bufferDriveRate).ToString());
                return (DriveRates)bufferDriveRate;
            }
            set
            {
                tl.LogMessage("TrackingRate Set", ((int)value).ToString());
                if (TrackingRate == value)
                    return;
                lock (buffLock)
                {
                    bufferDriveRate = value;
                }
                //is tracking, automatically apply changes
                if (Tracking) Tracking = true;
            }
        }

        public ITrackingRates TrackingRates
        {
            get
            {
                ITrackingRates trackingRates = new TrackingRates();
                tl.LogMessage("TrackingRates", "Get - ");
                foreach (DriveRates driveRate in trackingRates)
                {
                    tl.LogMessage("TrackingRates", "Get - " + driveRate.ToString());
                }
                return trackingRates;
            }
        }

        public DateTime UTCDate
        {
            get
            {
                CheckConnected("UTCDate Get");
                var raw = string.Format("{0} {1}", CommandString("GC"), CommandString("GL"));
                var offset = Convert.ToInt32(CommandString("GG"));
                var r = DateTime.ParseExact(raw, "MM/dd/yy HH:mm:ss", CultureInfo.InvariantCulture).AddHours(offset);
                tl.LogMessage("TrackingRates", "Get - " + raw);
                return r;
            }
            set
            {
                tl.LogMessage("UTCDate Set", String.Format("MM/dd/yy HH:mm:ss", value));
                value.AddHours(Convert.ToInt32(CommandString("GG")) * -1);
                CommandString("SC " + value.ToString("MM/dd/yy", CultureInfo.InvariantCulture));
                CommandString("SL " + value.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
            }
        }

        public void Unpark()
        {
            CheckConnected("Unpark");
            if (!AtPark)
                return;
            CommandBlind("YL");
        }

        #endregion

        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister)
        {
            using (var P = new ASCOM.Utilities.Profile())
            {
                P.DeviceType = "Telescope";
                if (bRegister)
                {
                    P.Register(driverID, driverDescription);
                }
                else
                {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t)
        {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t)
        {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected
        {
            get
            {
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Telescope";
                traceState = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
                legacyPulseGuide = Convert.ToBoolean(driverProfile.GetValue(driverID, legacyPulseGuideProfileName, string.Empty, legacyPulseGuideDefault));
                comPort = driverProfile.GetValue(driverID, comPortProfileName, string.Empty, comPortDefault);
                commandTimeout = Convert.ToInt32(driverProfile.GetValue(driverID, commandTimeoutProfileName, string.Empty, commandTimeoutDefault.ToString()));
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Telescope";
                driverProfile.WriteValue(driverID, traceStateProfileName, traceState.ToString());
                driverProfile.WriteValue(driverID, legacyPulseGuideProfileName, legacyPulseGuide.ToString());
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
                driverProfile.WriteValue(driverID, commandTimeoutProfileName, commandTimeout.ToString());
            }
        }

        #endregion

    }
}
