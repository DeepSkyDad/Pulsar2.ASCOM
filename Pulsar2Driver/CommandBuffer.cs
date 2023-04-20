using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ASCOM.Pulsar2
{
    public class CommandBuffer
    {
        class Command
        {
            public Command(int lifeTime)
            {
                Lifetime = lifeTime;
            }

            public int Lifetime { get; set; }
            public string Value { get; set; }
            public DateTime ValueDT { get; set; }
        }

        private Dictionary<string, Command> _buffer = new Dictionary<string, Command>() {
            { Helper.FormatCommand("GA"), new Command(3000) }, //altitude
            { Helper.FormatCommand("GZ"), new Command(3000) }, //azimuth
            { Helper.FormatCommand("GR"), new Command(3000) }, //RA
            { Helper.FormatCommand("GD"), new Command(3000) }, //DEC
            { Helper.FormatCommand("GS"), new Command(3000) }, //sidereal time
            { Helper.FormatCommand("YGi"), new Command(10000) }, //Is slewing
            { Helper.FormatCommand("YGk"), new Command(10000) }, //Is parked
            { Helper.FormatCommand("YGZ"), new Command(10000) }, //RA/DEC rate
            { Helper.FormatCommand("Gt"), new Command(60000) }, //latitude
            { Helper.FormatCommand("Gg"), new Command(60000) }, //longitude
            { Helper.FormatCommand("Gd"), new Command(10000) }, //target declination
            { Helper.FormatCommand("Gr"), new Command(10000) }, //target ra
            { Helper.FormatCommand("YGS"), new Command(10000) }, //tracking
            { Helper.FormatCommand("YGW"), new Command(10000) } //side of pier
        };

        private Dictionary<string, List<string>> _clearBufferMap = new Dictionary<string, List<string>>() {
            { "#:CM", new List<string>() { Helper.FormatCommand("GR"), Helper.FormatCommand("GD") } },
            { "#:MS", new List<string>() { Helper.FormatCommand("YGi") } },
            { "#:Q", new List<string>() { Helper.FormatCommand("YGi") } },
            { "#:YH", new List<string>() { Helper.FormatCommand("YGk") } },
            { "#:YL", new List<string>() { Helper.FormatCommand("YGk") } },
            { "#:St", new List<string>() { Helper.FormatCommand("Gt") } },
            { "#:Sg", new List<string>() { Helper.FormatCommand("Gg") } },
            { "#:Sd", new List<string>() { Helper.FormatCommand("Gd") } },
            { "#:Sr", new List<string>() { Helper.FormatCommand("Gr") } },
            { "#:YSS", new List<string>() { Helper.FormatCommand("YGS") } },
            { "#:YSZ", new List<string>() { Helper.FormatCommand("YGZ") } }
        };

        public string GetBufferValue(string command)
        {
            if (!_buffer.ContainsKey(command))
                return null;

            var bufferCmd = _buffer[command];
            if ((DateTime.Now - bufferCmd.ValueDT).TotalMilliseconds > bufferCmd.Lifetime)
            {
                return null;
            }

            return bufferCmd.Value;
        }

        public void ClearBufferValues(string command)
        {
            foreach (var kv in _clearBufferMap)
            {
                if (command.StartsWith(kv.Key))
                {
                    foreach (var clearCmd in kv.Value)
                    {
                        _buffer[clearCmd].ValueDT = DateTime.Now.AddDays(-1);
                    }
                }

            }   
        }

        public bool SetBufferValue(string command, string value)
        {
            if (!_buffer.ContainsKey(command))
                return false;

            var bufferCmd = _buffer[command];
            bufferCmd.Value = value;
            bufferCmd.ValueDT = DateTime.Now;
            return true;
        }
 
    }
}
