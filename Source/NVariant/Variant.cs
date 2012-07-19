// Copyright (C) Mark Alan Jones 2011
// This code is published under the Microsoft Public License (Ms-Pl)
// A copy of the Ms-Pl license is included with the source and
// binary distributions or online at
// http://www.microsoft.com/opensource/licenses.mspx#Ms-PL
//
// Variant.cs

using System;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

namespace NVariant
{
    /// <summary>
    /// A variant is ValueType capable of holding any of the basic CLR data
    /// types.
    /// <para>
    /// It provides methods for converting between compatible types.
    /// A union of the basic types is emulated by an explicit structure
    /// layout to avoid boxing of simple values.
    /// </para>
    /// </summary>
    [Serializable]
    public struct Variant : ISerializable, IEquatable<Variant>
    {
        private string strObj;
        private Union union;

        public Variant(SerializationInfo info, StreamingContext context)
            : this()
        {
            TypeCode = (TypeCode)info.GetValue("TypeCode", typeof(TypeCode));
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    union.Boolean = (bool)info.GetValue("Value", typeof(bool));
                    break;

                case TypeCode.Char:
                    union.Char = (char)info.GetValue("Value", typeof(char));
                    break;

                case TypeCode.SByte:
                    union.SByte = (sbyte)info.GetValue("Value", typeof(sbyte));
                    break;

                case TypeCode.Byte:
                    union.Byte = (byte)info.GetValue("Value", typeof(byte));
                    break;

                case TypeCode.Int16:
                    union.Int16 = (short)info.GetValue("Value", typeof(short));
                    break;

                case TypeCode.UInt16:
                    union.UInt16 = (ushort)info.GetValue("Value", typeof(ushort));
                    break;

                case TypeCode.Int32:
                    union.Int32 = (int)info.GetValue("Value", typeof(int));
                    break;

                case TypeCode.UInt32:
                    union.UInt32 = (uint)info.GetValue("Value", typeof(uint));
                    break;

                case TypeCode.Int64:
                    union.Int64 = (long)info.GetValue("Value", typeof(long));
                    break;

                case TypeCode.UInt64:
                    union.UInt64 = (ulong)info.GetValue("Value", typeof(ulong));
                    break;

                case TypeCode.Single:
                    union.Single = (float)info.GetValue("Value", typeof(float));
                    break;

                case TypeCode.Double:
                    union.Double = (double)info.GetValue("Value", typeof(double));
                    break;

                case TypeCode.Decimal:
                    union.Decimal = (decimal)info.GetValue("Value", typeof(decimal));
                    break;

                case TypeCode.DateTime:
                    union.DateTime = (DateTime)info.GetValue("Value", typeof(DateTime));
                    break;

                case TypeCode.String:
                    strObj = (string)info.GetValue("Value", typeof(string));
                    break;
            }
        }

        /// <summary>
        /// Constructs a variant from a Boolean
        /// </summary>
        /// <param name="value"></param>
        public Variant(bool value)
            : this()
        {
            union = new Union { Boolean = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.Byte
        /// </summary>
        /// <param name="value"></param>
        public Variant(byte value)
            : this()
        {
            union = new Union { Byte = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from an System.SByte
        /// </summary>
        /// <param name="value"></param>
        [CLSCompliant(false)]
        public Variant(sbyte value)
            : this()
        {
            union = new Union { SByte = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.Int16
        /// </summary>
        /// <param name="value"></param>
        public Variant(short value)
            : this()
        {
            union = new Union { Int16 = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.UInt16
        /// </summary>
        /// <param name="value"></param>
        [CLSCompliant(false)]
        public Variant(ushort value)
            : this()
        {
            union = new Union { UInt16 = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.Int32
        /// </summary>
        /// <param name="value"></param>
        public Variant(int value)
            : this()
        {
            union = new Union { Int32 = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.UInt32
        /// </summary>
        /// <param name="value"></param>
        [CLSCompliant(false)]
        public Variant(uint value)
            : this()
        {
            union = new Union { UInt32 = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.Int64
        /// </summary>
        /// <param name="value"></param>
        public Variant(long value)
            : this()
        {
            union = new Union { Int64 = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.UInt64
        /// </summary>
        /// <param name="value"></param>
        [CLSCompliant(false)]
        public Variant(ulong value)
            : this()
        {
            union = new Union { UInt64 = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.Single
        /// </summary>
        /// <param name="value"></param>
        public Variant(float value)
            : this()
        {
            union = new Union { Single = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.Double
        /// </summary>
        /// <param name="value"></param>
        public Variant(double value)
            : this()
        {
            union = new Union { Double = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.Decimal
        /// </summary>
        /// <param name="value"></param>
        public Variant(decimal value)
            : this()
        {
            union = new Union { Decimal = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.DateTime
        /// </summary>
        /// <param name="value"></param>
        public Variant(DateTime value)
            : this()
        {
            union = new Union { DateTime = value };
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.String
        /// </summary>
        /// <param name="value"></param>
        public Variant(string value)
            : this()
        {
            strObj = value;
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// Constructs a variant from a System.Char
        /// </summary>
        /// <param name="value"></param>
        public Variant(char value)
            : this()
        {
            union.Char = value;
            TypeCode = value.GetTypeCode();
        }

        /// <summary>
        /// True if the variant is empty
        /// </summary>
        public bool IsEmpty
        {
            get { return TypeCode == TypeCode.Empty; }
        }

        /// <summary>
        /// TypeCode of the variant value
        /// </summary>
        public TypeCode TypeCode { get; private set; }

        #region IEquatable<Variant> Members

        /// <summary>
        /// Tests another variant against this variant for equality
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(Variant other)
        {
            return ToString().Equals(other.ToString());
        }

        #endregion IEquatable<Variant> Members

        #region ISerializable Members

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("TypeCode", TypeCode);
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    info.AddValue("Value", union.Boolean);
                    break;

                case TypeCode.Char:
                    info.AddValue("Value", union.Char);
                    break;

                case TypeCode.SByte:
                    info.AddValue("Value", union.SByte);
                    break;

                case TypeCode.Byte:
                    info.AddValue("Value", union.Byte);
                    break;

                case TypeCode.Int16:
                    info.AddValue("Value", union.Int16);
                    break;

                case TypeCode.UInt16:
                    info.AddValue("Value", union.UInt16);
                    break;

                case TypeCode.Int32:
                    info.AddValue("Value", union.Int32);
                    break;

                case TypeCode.UInt32:
                    info.AddValue("Value", union.UInt32);
                    break;

                case TypeCode.Int64:
                    info.AddValue("Value", union.Int64);
                    break;

                case TypeCode.UInt64:
                    info.AddValue("Value", union.UInt64);
                    break;

                case TypeCode.Single:
                    info.AddValue("Value", union.Single);
                    break;

                case TypeCode.Double:
                    info.AddValue("Value", union.Double);
                    break;

                case TypeCode.Decimal:
                    info.AddValue("Value", union.Decimal);
                    break;

                case TypeCode.DateTime:
                    info.AddValue("Value", union.DateTime);
                    break;

                case TypeCode.String:
                    info.AddValue("Value", strObj);
                    break;
            }
        }

        #endregion ISerializable Members

        public static bool operator !=(Variant v1, Variant v2)
        {
            return !(v1 == v2);
        }

        public static bool operator ==(Variant v1, Variant v2)
        {
            return v1.Equals(v2);
        }

        /// <summary>
        /// Tests an object against the variant for equality
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
            {
                return false;
            }
            return obj is Variant && Equals((Variant)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + ToString().GetHashCode();
                hash = hash * 23 + TypeCode.GetHashCode();
                return hash;
            }
        }

        /// <summary>
        /// Converts the variant value to a System.Boolean
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public bool ToBoolean()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return union.Boolean;
                case TypeCode.SByte:
                    return Convert.ToBoolean(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToBoolean(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToBoolean(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToBoolean(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToBoolean(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToBoolean(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToBoolean(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToBoolean(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToBoolean(union.Single);
                case TypeCode.Double:
                    return Convert.ToBoolean(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToBoolean(union.Decimal);
                case TypeCode.String:
                    return Convert.ToBoolean(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.Byte
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public byte ToByte()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToByte(union.Boolean);
                case TypeCode.Char:
                    return Convert.ToByte(union.Char);
                case TypeCode.SByte:
                    return Convert.ToByte(union.SByte);
                case TypeCode.Byte:
                    return union.Byte;
                case TypeCode.Int16:
                    return Convert.ToByte(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToByte(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToByte(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToByte(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToByte(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToByte(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToByte(union.Single);
                case TypeCode.Double:
                    return Convert.ToByte(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToByte(union.Decimal);
                case TypeCode.String:
                    return Convert.ToByte(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.Char
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public char ToChar()
        {
            switch (TypeCode)
            {
                case TypeCode.Char:
                    return union.Char;
                case TypeCode.SByte:
                    return Convert.ToChar(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToChar(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToChar(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToChar(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToChar(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToChar(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToChar(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToChar(union.UInt64);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.DateTime
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public DateTime ToDateTime()
        {
            switch (TypeCode)
            {
                case TypeCode.DateTime:
                    return union.DateTime;
                case TypeCode.String:
                    return Convert.ToDateTime(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.Decimal
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public Decimal ToDecimal()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToDecimal(union.Boolean);
                case TypeCode.SByte:
                    return Convert.ToDecimal(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToDecimal(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToDecimal(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToUInt16(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToDecimal(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToDecimal(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToDecimal(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToDecimal(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToDecimal(union.Single);
                case TypeCode.Double:
                    return Convert.ToDecimal(union.Double);
                case TypeCode.Decimal:
                    return union.Decimal;
                case TypeCode.String:
                    return Convert.ToDecimal(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.Double
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public double ToDouble()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToDouble(union.Boolean);
                case TypeCode.SByte:
                    return Convert.ToDouble(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToDouble(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToDouble(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToDouble(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToDouble(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToDouble(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToDouble(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToDouble(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToDouble(union.Single);
                case TypeCode.Double:
                    return union.Double;
                case TypeCode.Decimal:
                    return Convert.ToDouble(union.Decimal);
                case TypeCode.String:
                    return Convert.ToDouble(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.Int16
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public short ToInt16()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToInt16(union.Boolean);
                case TypeCode.Char:
                    return Convert.ToInt16(union.Char);
                case TypeCode.SByte:
                    return Convert.ToInt16(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToInt16(union.Byte);
                case TypeCode.Int16:
                    return union.Int16;
                case TypeCode.UInt16:
                    return Convert.ToInt16(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToInt16(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToInt16(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToInt16(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToInt16(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToInt16(union.Single);
                case TypeCode.Double:
                    return Convert.ToInt16(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToInt16(union.Decimal);
                case TypeCode.String:
                    return Convert.ToInt16(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.Int32
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public int ToInt32()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToInt32(union.Boolean);
                case TypeCode.Char:
                    return Convert.ToInt32(union.Char);
                case TypeCode.SByte:
                    return Convert.ToInt32(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToInt32(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToInt32(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToInt32(union.UInt16);
                case TypeCode.Int32:
                    return union.Int32;
                case TypeCode.UInt32:
                    return Convert.ToInt32(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToInt32(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToInt32(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToInt32(union.Single);
                case TypeCode.Double:
                    return Convert.ToInt32(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToInt32(union.Decimal);
                case TypeCode.String:
                    return Convert.ToInt32(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.Int64
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public long ToInt64()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToInt64(union.Boolean);
                case TypeCode.Char:
                    return Convert.ToInt64(union.Char);
                case TypeCode.SByte:
                    return Convert.ToInt64(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToInt64(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToInt64(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToInt64(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToInt64(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToInt64(union.UInt32);
                case TypeCode.Int64:
                    return union.Int64;
                case TypeCode.UInt64:
                    return Convert.ToInt64(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToInt64(union.Single);
                case TypeCode.Double:
                    return Convert.ToInt64(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToInt64(union.Decimal);
                case TypeCode.String:
                    return Convert.ToInt64(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.SByte
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        [CLSCompliant(false)]
        public sbyte ToSByte()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToSByte(union.Boolean);
                case TypeCode.Char:
                    return Convert.ToSByte(union.Char);
                case TypeCode.SByte:
                    return union.SByte;
                case TypeCode.Byte:
                    return Convert.ToSByte(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToSByte(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToSByte(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToSByte(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToSByte(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToSByte(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToSByte(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToSByte(union.Single);
                case TypeCode.Double:
                    return Convert.ToSByte(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToSByte(union.Decimal);
                case TypeCode.String:
                    return Convert.ToSByte(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.Single
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public float ToSingle()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToSingle(union.Boolean);
                case TypeCode.SByte:
                    return Convert.ToSingle(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToSingle(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToSingle(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToSingle(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToSingle(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToSingle(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToSingle(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToSingle(union.UInt64);
                case TypeCode.Single:
                    return union.Single;
                case TypeCode.Double:
                    return Convert.ToSingle(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToSingle(union.Decimal);
                case TypeCode.String:
                    return Convert.ToSingle(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.String
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        public override string ToString()
        {
            // our struct is immutable so cache the string
            return strObj ?? (strObj = InternalToString());
        }

        /// <summary>
        /// Converts the variant value to a System.UInt16
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        [CLSCompliant(false)]
        public ushort ToUInt16()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToUInt16(union.Boolean);
                case TypeCode.Char:
                    return Convert.ToUInt16(union.Char);
                case TypeCode.SByte:
                    return Convert.ToUInt16(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToUInt16(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToUInt16(union.Int16);
                case TypeCode.UInt16:
                    return union.UInt16;
                case TypeCode.Int32:
                    return Convert.ToUInt16(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToUInt16(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToUInt16(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToUInt16(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToUInt16(union.Single);
                case TypeCode.Double:
                    return Convert.ToUInt16(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToUInt16(union.Decimal);
                case TypeCode.String:
                    return Convert.ToUInt16(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.UInt32
        /// </summary>        
        /// <exception cref="InvalidCastException"></exception>
        [CLSCompliant(false)]
        public uint ToUInt32()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToUInt32(union.Boolean);
                case TypeCode.Char:
                    return Convert.ToUInt32(union.Char);
                case TypeCode.SByte:
                    return Convert.ToUInt32(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToUInt32(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToUInt32(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToUInt32(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToUInt32(union.Int32);
                case TypeCode.UInt32:
                    return union.UInt32;
                case TypeCode.Int64:
                    return Convert.ToUInt32(union.Int64);
                case TypeCode.UInt64:
                    return Convert.ToUInt32(union.UInt64);
                case TypeCode.Single:
                    return Convert.ToUInt32(union.Single);
                case TypeCode.Double:
                    return Convert.ToUInt32(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToUInt32(union.Decimal);
                case TypeCode.String:
                    return Convert.ToUInt32(strObj);
            }
            throw new InvalidCastException();
        }

        /// <summary>
        /// Converts the variant value to a System.UInt64
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        [CLSCompliant(false)]
        public ulong ToUInt64()
        {
            switch (TypeCode)
            {
                case TypeCode.Boolean:
                    return Convert.ToUInt64(union.Boolean);
                case TypeCode.Char:
                    return Convert.ToUInt64(union.Char);
                case TypeCode.SByte:
                    return Convert.ToUInt64(union.SByte);
                case TypeCode.Byte:
                    return Convert.ToUInt64(union.Byte);
                case TypeCode.Int16:
                    return Convert.ToUInt64(union.Int16);
                case TypeCode.UInt16:
                    return Convert.ToUInt64(union.UInt16);
                case TypeCode.Int32:
                    return Convert.ToUInt64(union.Int32);
                case TypeCode.UInt32:
                    return Convert.ToUInt64(union.UInt32);
                case TypeCode.Int64:
                    return Convert.ToUInt64(union.Int64);
                case TypeCode.UInt64:
                    return union.UInt64;
                case TypeCode.Single:
                    return Convert.ToUInt64(union.Single);
                case TypeCode.Double:
                    return Convert.ToUInt64(union.Double);
                case TypeCode.Decimal:
                    return Convert.ToUInt64(union.Decimal);
                case TypeCode.String:
                    return Convert.ToUInt64(strObj);
            }
            throw new InvalidCastException();
        }

        private string InternalToString()
        {
            switch (TypeCode)
            {
                case TypeCode.Empty:
                    return string.Empty;
                case TypeCode.Boolean:
                    return union.Boolean.ToString();
                case TypeCode.Char:
                    return union.Char.ToString();
                case TypeCode.SByte:
                    return union.SByte.ToString();
                case TypeCode.Byte:
                    return union.Byte.ToString();
                case TypeCode.Int16:
                    return union.Int16.ToString();
                case TypeCode.UInt16:
                    return union.UInt16.ToString();
                case TypeCode.Int32:
                    return union.Int32.ToString();
                case TypeCode.UInt32:
                    return union.UInt32.ToString();
                case TypeCode.Int64:
                    return union.Int64.ToString();
                case TypeCode.UInt64:
                    return union.UInt64.ToString();
                case TypeCode.Single:
                    return union.Single.ToString();
                case TypeCode.Double:
                    return union.Double.ToString();
                case TypeCode.Decimal:
                    return union.Decimal.ToString();
                case TypeCode.DateTime:
                    return union.DateTime.ToString();
            }
            throw new InvalidCastException();
        }

        #region Implicit casts

        public static implicit operator Variant(bool value)
        {
            return new Variant(value);
        }

        public static implicit operator Variant(byte value)
        {
            return new Variant(value);
        }

        [CLSCompliant(false)]
        public static implicit operator Variant(sbyte value)
        {
            return new Variant(value);
        }

        [CLSCompliant(false)]
        public static implicit operator Variant(ushort value)
        {
            return new Variant(value);
        }

        public static implicit operator Variant(short value)
        {
            return new Variant(value);
        }

        [CLSCompliant(false)]
        public static implicit operator Variant(uint value)
        {
            return new Variant(value);
        }

        public static implicit operator Variant(int value)
        {
            return new Variant(value);
        }

        public static implicit operator Variant(long value)
        {
            return new Variant(value);
        }

        [CLSCompliant(false)]
        public static implicit operator Variant(ulong value)
        {
            return new Variant(value);
        }

        public static implicit operator Variant(float value)
        {
            return new Variant(value);
        }

        public static implicit operator Variant(double value)
        {
            return new Variant(value);
        }

        public static implicit operator Variant(decimal value)
        {
            return new Variant(value);
        }

        public static implicit operator Variant(DateTime value)
        {
            return new Variant(value);
        }

        public static implicit operator Variant(string value)
        {
            return new Variant(value);
        }

        #endregion Implicit casts

        #region Nested type: Union

        [StructLayout(LayoutKind.Explicit)]
        private struct Union
        {
            [FieldOffset(0)]
            public bool Boolean;

            [FieldOffset(0)]
            public byte Byte;

            [FieldOffset(0)]
            public char Char;

            [FieldOffset(0)]
            public DateTime DateTime;

            [FieldOffset(0)]
            public decimal Decimal;

            [FieldOffset(0)]
            public double Double;

            [FieldOffset(0)]
            public short Int16;

            [FieldOffset(0)]
            public int Int32;

            [FieldOffset(0)]
            public long Int64;

            [FieldOffset(0)]
            public sbyte SByte;

            [FieldOffset(0)]
            public float Single;

            [FieldOffset(0)]
            public ushort UInt16;

            [FieldOffset(0)]
            public uint UInt32;

            [FieldOffset(0)]
            public ulong UInt64;
        }

        #endregion Nested type: Union
    }
}