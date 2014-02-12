using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace Tabbles.OutlookAddIn.Common.Types
{
    [Serializable()]
    public abstract class Icon : ISerializable
    {
        public Icon()
        {
        }

        public Icon(SerializationInfo info, StreamingContext context)
        {
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
        }
    }

    [Serializable()]
    public class IconOther : Icon
    {
        public IconOther() : base()
        {
        }

        public IconOther(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public new void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }

    [Serializable()]
    public class Generic : ISerializable
    {
        public string name { get; set; }
        public string commandLine { get; set; }
        public Icon icon { get; set; }
        public bool showCommandLine { get; set; }

        public Generic()
        {
        }

        public Generic(SerializationInfo info, StreamingContext context)
        {
            name = info.GetString("name");
            commandLine = info.GetString("commandLine");
            icon = info.GetValue("icon", typeof(Icon)) as Icon;
            showCommandLine = info.GetBoolean("showCommandLine");
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("name", name);
            info.AddValue("commandLine", commandLine);
            info.AddValue("icon", icon);
            info.AddValue("showCommandLine", showCommandLine);
        }
    }
}
