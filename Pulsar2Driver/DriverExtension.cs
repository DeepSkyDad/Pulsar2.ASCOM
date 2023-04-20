using ASCOM.DeviceInterface;
using ASCOM.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace ASCOM.Pulsar2
{
    public partial class Telescope
    {
        //private const long commandTimeoutMs = 130; //timeout, defined by Andras Dan (maximum pooling frequency is (1000/timeout)Hz)
        private Version minimumFirmwareVerson = new Version("5.60");

        //for locking serial communication
        private Mutex serialMutex = null;
        private bool createdNew;


        private void CheckVersion()
        {
            CheckConnected("CheckVersion");
            var currentVersion = new Version(Regex.Matches(CommandString("YV"), "PULSAR V([\\d.]+)\\w.*")[0].Groups[1].Value);
            if(minimumFirmwareVerson > currentVersion)
                throw new DriverException(string.Format("Please upgrade firmware before using {0} driver (minimum version: {1})", driverID, minimumFirmwareVerson.ToString()));
        }

        //commands buffer
        private CommandBuffer _cmdBuffer;

        private void CommandBlind(string command)
        {
            CommandBlind(command, false);
        }

        private bool CommandBool(string command)
        {
            return CommandBool(command, false);
        }

        private string CommandString(string command)
        {
            return CommandString(command, false);
        }

        /// <summary>
        /// it's a good idea to put all the low level communication with the device here - then all communication calls this function.
        /// </summary>
        /// <param name="command">given command</param>
        /// <param name="raw">if true, skip formatting. If false, format command text</param>
        /// <param name="async">if true, do not wait for response</param>
        /// <returns>response text, in case of empty command or async empty string</returns>
        private string CommandString(string command, bool raw, bool async)
        {
            bool hasHandle = false;

            //init
            if(serialMutex == null && _cmdBuffer == null) {
                //set up security for multi-user usage + add support for localized systems (don't use just "Everyone") 
                MutexAccessRule allowEveryoneRule = new MutexAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), MutexRights.FullControl, AccessControlType.Allow);
                MutexSecurity securitySettings = new MutexSecurity();
                securitySettings.AddAccessRule(allowEveryoneRule);
                serialMutex = new Mutex(false, driverID, out createdNew, securitySettings);

                _cmdBuffer = new CommandBuffer();
            }


            var response = string.Empty;

            //if command is empty, return immediately
            if (string.IsNullOrWhiteSpace(command))
                return response;

            //format command if not raw
            if (!raw)
                command = Helper.FormatCommand(command);

            try
            {
                try
                {
                    //mutex lock to ensure that only one command is in progress at a time (http://www.ascom-standards.org/Help/Developer/html/caf6b21d-755a-4f1c-891f-ce971a9a2f79.htm)
                    //note, we want to time out here instead of waiting forever
                    hasHandle = serialMutex.WaitOne(commandTimeout, false);
                    if (hasHandle == false)
                        throw new TimeoutException("Timeout waiting for exclusive access");
                    //tl.LogMessage("CommandString mutex", "mutex acquired");
                }
                catch (AbandonedMutexException)
                {
                    //log the fact that the mutex was abandoned in another process, it will still get acquired
                    //tl.LogMessage("CommandString mutex", "mutex acquired via abandon");
                    hasHandle = true;
                }

                //serial communication
                var watch = Stopwatch.StartNew();

                if (async)
                {
                    _cmdBuffer.ClearBufferValues(command);
                    tl.LogMessage("CommandString async", string.Format("Sending command {0}", command));
                    tls.LogMessage("Request async", command);
                    serial.ClearBuffers();
                    serial.Transmit(command); //async message - do not wait for response

                    //handle Pulsar2 bug -> wait 1s after YSS - should be fixed in 5.58a
                    if (command.StartsWith("#:YSS"))
                    {
                        utilities.WaitForMilliseconds(1000);
                    }
                }
                else
                {
                    //check if result is in buffer
                    tls.LogMessage("Request", command);
                    if (_cmdBuffer.GetBufferValue(command) != null)
                    {
                        response = _cmdBuffer.GetBufferValue(command).TrimEnd('#');
                        tls.LogMessage("Response", response + " (fetched from buffer!)");
                        tl.LogMessage("CommandString sync", string.Format("Response for {0} found in buffer: {1}", command, response));
                        return response;
                    }
                    else
                    {
                        serial.ClearBuffers();
                        _cmdBuffer.ClearBufferValues(command);
                        serial.Transmit(command);
                        response = serial.ReceiveTerminated("#"); //wait until termination character

                        var isBuffered = _cmdBuffer.SetBufferValue(command, response);
                        tls.LogMessage("Response", response + (isBuffered ? " (saved to buffer)" : string.Empty));
                        tl.LogMessage("CommandString sync", string.Format("Response for {0} received: {1} {2}", command, response, isBuffered ? "(saved to buffer)" : string.Empty));

                        //hack requested by Andras -> MS sets YGi to 1 in buffer because firmware bug (returns 0 immediately after MS) this was fixed in 5.60
                        //if (command == "#:MS#")
                        //{
                        //    _cmdBuffer.SetBufferValue("#:YGi#", "1#");
                        //}

                        //handler response bugs - sometimes leading 1 is added... - should be fixed in 5.58a
                        //if (command == "#:YSS#")
                        //{
                        //    //trim precedding spam characters (for example some responses are like this "10,0", "110,0")
                        //    response = Regex.Matches(response, "(\\d,\\d)$")[0].Groups[1].Value;
                        //}
                        
                    }

                    response = response.TrimEnd('#');
                }

                //maximum frequency is XYHz so execution must take at least XYms
                //watch.Stop();
                //if (watch.ElapsedMilliseconds < commandTimeoutMs)
                //utilities.WaitForMilliseconds((int)(commandTimeoutMs - watch.ElapsedMilliseconds));
            }
            catch (COMException ce)
            {
                try
                {
                    //HACK requested by Andras: retry if timeout occurs due to Pulsar bugs
                    if (ce.Message.StartsWith("Timed"))
                    {
                        tl.LogMessage("CommandString sync retry", string.Format("Timeout occured for command {0} - retry", command));
                        serial.Transmit(command);
                        response = serial.ReceiveTerminated("#"); //wait until termination character

                        var isBuffered = _cmdBuffer.SetBufferValue(command, response);
                        tls.LogMessage("Response", response + (isBuffered ? " (saved to buffer)" : string.Empty));
                        tl.LogMessage("CommandString sync retry", string.Format("Response for {0} received: {1} {2}", command, response, isBuffered ? "(saved to buffer)" : string.Empty));
                    }
                }
                catch (Exception e)
                {
                    tl.LogMessage("CommandString sync retry error", string.Format("Command: {0}, Message: {1}, StackTrace: {2}", command, e.Message, e.StackTrace));
                    throw;
                }
            }
            catch (Exception e)
            {
                tl.LogMessage("CommandString error", string.Format("Command: {0}, Message: {1}, StackTrace: {2}", command, e.Message, e.StackTrace));
                throw;
            }
            finally
            {
                if (hasHandle)
                {
                    serialMutex.ReleaseMutex();
                    //tl.LogMessage("CommandString mutex", "mutex released");
                }
            }

            return response;
        }

        private bool IsMountParking
        {
            get
            {
                CheckConnected("IsMountParking");
                var r = CommandBool("YGj");
                tl.LogMessage("IsMountParking", "Get - " + r.ToString());
                return r;
            }
        }

        private string FormatYSZRadMinDouble(double value)
        {
            var result = value.ToString("N7", CultureInfo.InvariantCulture);
            if (result[0] != '-')
                result = "+" + result;
            return result;
        }

        private string[] ParseYGS(string result)
        {
            //requested by Andras - ygs sometimes returns "1" before RA rate which is wrong (e.g. "1+0.0000000,-0.0000308" instead of "+0.0000000,-0.0000308")
            var matches = Regex.Matches(result, "\\d*([+-][\\d.]+),([+-][\\d.]+)")[0];
            var raRate = matches.Groups[1].Value;
            var decRate = matches.Groups[2].Value;
             tl.LogMessage("YGS Parse", string.Format("input: {0} output: {1},{2}", result, raRate, decRate));
            return new string[2] { raRate, decRate };
        }

        private double[] ParseYGSDouble(string result)
        {

            var valuesStr = ParseYGS(result);
            double raRate, decRate;
            string raRateStr = valuesStr[0];
            string decRateStr = valuesStr[1];

            var nfi = new NumberFormatInfo { CurrencyDecimalSeparator = "." };
            double.TryParse(raRateStr, NumberStyles.Currency, nfi, out raRate);
            double.TryParse(decRateStr, NumberStyles.Currency, nfi, out decRate);

            return new double[2] { raRate, decRate };
        }

        //TODO: on hold for now - when time comes, round given input rate to closest step (confirmed by Andras)
        //private void SetPulseGuideRate(double rate)
        //{
        //    //by Andras - in Pulsar2 both RA and DEC guiding rate have same value so treat them as one (this method is called by GuideRateRightAscention and GuideRateDeclination)
        //    var rateRange = AxisRates(TelescopeAxes.axisPrimary)[0];
        //    //supported guiding rates are 10%-90% of sidereal with 10% step
        //    if (Math.Abs(rate) < (0.1*<sidereal>) || Math.Abs(rate) > (<sidereal>*0.9))
        //        throw new DriverException(string.Format("Invalid rate - {0}", rate));


        //    //round given rate to 10% accuracy,  (pulsar2 guide rate supports 10%-90% with 10% step)
        //    int percantage = (int)(Math.Round(rate/rateRange.Maximum*100)/10);

        //    CommandBlind("YSA")
        //}

    }
}
