using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace Tabbles.OutlookAddIn.Common.Messages
{
    #region from tabbles to outlook
    [Serializable()]
    public class MsgAtomKeysDeleted : ISerializable
    {
        public IEnumerable<string> AtomKeys { get; set; }

        public MsgAtomKeysDeleted()
        {
        }

        public MsgAtomKeysDeleted(SerializationInfo info, StreamingContext context)
        {
            AtomKeys = info.GetValue("AtomKeys", typeof(IEnumerable<string>)) as IEnumerable<string>;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("AtomKeys", AtomKeys);
        }
    }

    [Serializable()]
    public class MsgAtomKeyCreated : ISerializable
    {
        public string AtomKeyName { get; set; }
        public string AtomKeyColor { get; set; }

        public MsgAtomKeyCreated()
        {
        }

        public MsgAtomKeyCreated(SerializationInfo info, StreamingContext context)
        {
            AtomKeyName = info.GetString("AtomKeyName");
            AtomKeyColor = info.GetString("AtomKeyColor");
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("AtomKeyName", AtomKeyName);
            info.AddValue("AtomKeyColor", AtomKeyColor);
        }
    }

    [Serializable()]
    public class NameColorPair : ISerializable
    {
        public string Name { get; set; }
        public string Color { get; set; }

        public NameColorPair()
        {
        }

        public NameColorPair(SerializationInfo info, StreamingContext context)
        {
            Name = info.GetString("Name");
            Color = info.GetString("Color");
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("Name", Name);
            info.AddValue("Color", Color);
        }
    }


    [Serializable()]
    public class MsgGensTagged : ISerializable
    {
        public IEnumerable<string> gens { get; set; }
        public IEnumerable<NameColorPair> tags { get; set; }

        public MsgGensTagged()
        {
        }

        public MsgGensTagged(SerializationInfo info, StreamingContext context)
        {
            gens = info.GetValue("gens", typeof(IEnumerable<string>)) as IEnumerable<string>;
            tags = info.GetValue("tags", typeof(IEnumerable<NameColorPair>)) as IEnumerable<NameColorPair>;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("gens", gens);
            info.AddValue("tags", tags);
        }
    }

    [Serializable()]
    public class MsgGensUntagged : ISerializable
    {
        public IEnumerable<string> gens { get; set; }
        public IEnumerable<string> tags { get; set; }

        public MsgGensUntagged()
        {
        }

        public MsgGensUntagged(SerializationInfo info, StreamingContext context)
        {
            gens = info.GetValue("gens", typeof(IEnumerable<string>)) as IEnumerable<string>;
            tags = info.GetValue("tags", typeof(IEnumerable<string>)) as IEnumerable<string>;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("gens", gens);
            info.AddValue("tags", tags);
        }
    }

    [Serializable()]
    public class MsgAddTabblesToCategories : ISerializable
    {
        public IEnumerable<string> tags { get; set; }

        public MsgAddTabblesToCategories()
        {
        }

        public MsgAddTabblesToCategories(SerializationInfo info, StreamingContext context)
        {
            tags = info.GetValue("tags", typeof(IEnumerable<string>)) as IEnumerable<string>;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("tags", tags);
        }
    }

    [Serializable()]
    public class MsgRemoveTabblesFromCategories : ISerializable
    {
        public IEnumerable<string> tags { get; set; }

        public MsgRemoveTabblesFromCategories()
        {
        }

        public MsgRemoveTabblesFromCategories(SerializationInfo info, StreamingContext context)
        {
            tags = info.GetValue("tags", typeof(IEnumerable<string>)) as IEnumerable<string>;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("tags", tags);
        }
    }

    [Serializable()]
    public class MsgOpenMailsWithTags : ISerializable
    {
        public IEnumerable<string> tags { get; set; }

        public MsgOpenMailsWithTags()
        {
        }

        public MsgOpenMailsWithTags(SerializationInfo info, StreamingContext context)
        {
            tags = info.GetValue("tags", typeof(IEnumerable<string>)) as IEnumerable<string>;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("tags", tags);
        }
    }
    #endregion

    #region from outlook to tabbles
    //public class genericTaggedOrUntaggedWithGivenTabbles
    //{
    //    public string entryId { get; set; }
    //    public bool tagged { get; set; }
    //    public IEnumerable<string> tags { get; set; }
    //}

    [Serializable()]
    public class GenericChangedSomeCategory : ISerializable
    {
        //public string entryId { get; set; }

        public Types.Generic gen { get; set; }
        public Dictionary<string, string> categoriesWithColors { get; set; }

        public GenericChangedSomeCategory()
        {
        }

        public GenericChangedSomeCategory(SerializationInfo info, StreamingContext context)
        {
            gen = info.GetValue("gen", typeof(Types.Generic)) as Types.Generic;
            categoriesWithColors = info.GetValue("categoriesWithColors", typeof(Dictionary<string, string>))
                as Dictionary<string, string>;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("gen", gen);
            info.AddValue("categoriesWithColors", categoriesWithColors);
        }
    }


    [Serializable()]
    public class INeedToTagGenericsWithTabblesQuickTagDialog : ISerializable
    {
        public IEnumerable<Types.Generic> gens { get; set; }

        public INeedToTagGenericsWithTabblesQuickTagDialog()
        {
        }

        public INeedToTagGenericsWithTabblesQuickTagDialog(SerializationInfo info, StreamingContext context)
        {
            gens = info.GetValue("gens", typeof(IEnumerable<Types.Generic>)) as IEnumerable<Types.Generic>;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("gens", gens);
        }
    }

    [Serializable()]
    public class INeedToTagGenericWithKnownTags : ISerializable
    {
        public Types.Generic gen { get; set; }

        public IEnumerable<string> tags { get; set; }

        public INeedToTagGenericWithKnownTags()
        {
        }

        public INeedToTagGenericWithKnownTags(SerializationInfo info, StreamingContext context)
        {
            gen = info.GetValue("gen", typeof(Types.Generic)) as Types.Generic;
            tags = info.GetValue("tags", typeof(IEnumerable<string>)) as IEnumerable<string>;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("gen", gen);
            info.AddValue("tags", tags);
        }
    }

    [Serializable()]
    public class INeedToAddTabblesToCategories : ISerializable
    {
        public INeedToAddTabblesToCategories()
        {
        }

        public INeedToAddTabblesToCategories(SerializationInfo info, StreamingContext context)
        {
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
        }
    }

    [Serializable()]
    public class INeedToRemoveTabblesFromCategories : ISerializable
    {
        public INeedToRemoveTabblesFromCategories()
        {
        }

        public INeedToRemoveTabblesFromCategories(SerializationInfo info, StreamingContext context)
        {
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
        }
    }

    [Serializable()]
    public class INeedToPingTabbles : ISerializable
    {
        public INeedToPingTabbles()
        {
        }

        public INeedToPingTabbles(SerializationInfo info, StreamingContext context)
        {
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
        }
    }

    [Serializable()]
    public class INeedToOpenGenericInTabbles : ISerializable
    {
        public Types.Generic gen { get; set; }

        public INeedToOpenGenericInTabbles()
        {
        }

        public INeedToOpenGenericInTabbles(SerializationInfo info, StreamingContext context)
        {
            gen = info.GetValue("gen", typeof(Types.Generic)) as Types.Generic;
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("gen", gen);
        }
    }

    [Serializable()]
    public class INeedToOpenSearch : ISerializable
    {
        public INeedToOpenSearch()
        {
        }

        public INeedToOpenSearch(SerializationInfo info, StreamingContext context)
        {
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
        }
    }

    #endregion
}
