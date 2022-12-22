using System;

namespace Excel2Entity
{
    public class CsTypes
    {
        public CsTypes()
        {
        }

        public CsTypes(string name, Type value)
        {
            Name = name;
            Value = value;
        }

        /// <summary></summary>
        public string Name { get; set; }

        /// <summary></summary>
        public Type Value { get; set; }
    }
}
