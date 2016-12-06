[CmdletBinding(SupportsShouldProcess=$true)]
Param
(

    [Parameter(Mandatory=$True)]
    [string] $SourceGPOName,

	[Parameter()]
	[string] $TargetGPOName = $SourceGPOName,

    [Parameter()]
    [string] $SourceVersion = "15.0",

	[Parameter()]
	[string] $TargetVersion ="16.0"

)

function Copy-OfficeGPOSettings {
<#
.SYNOPSIS
Copies Group Policies between Office Versions. Defaults to: 15 (Office 2013) to 16 (Office 2016)
.DESCRIPTION
Given a source, target, and the filepath to C# support file, this cmdlet finds all the office 15 policies
in the source that are associated with the source and copies them over to the target as office 16 policies.
.PARAMETER SourceGPOName
The Name of the GPO that you wish to transfer office policies from. Defaults to 15.0 (Office 2013)
.PARAMETER TargetGPOName
The Name of the GPO that you wish to transfer office policies to. Defaults to 16.0 (Office 2016)
.PARAMETER SourceVersion
The version number of the office settings to copy
.PARAMETER TargetVersion
The version number of the office settings to set
.Example
./Copy-OfficeGPOSettings -SourceGPOName "Office Settings"
Default copy the office 15.0 (2013) policies within 'Office Settings' to office 16.0 (2016) policies within 'Office Settings'
.Example
./Copy-OfficeGPOSettings -SourceGPOName "Office Settings" -SourceVersion "14.0" -TargetVersion "15.0"
Copy the office 14.0 (2010) policies within 'Office Settings' to office 15.0 (2013) policies within 'Office Settings'
.Link
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts
#>
[CmdletBinding(SupportsShouldProcess=$true)]
Param
(

    [Parameter(Mandatory=$True)]
    [string] $SourceGPOName,

	[Parameter()]
	[string] $TargetGPOName = $SourceGPOName,

    [Parameter()]
    [string] $SourceVersion = "15.0",

	[Parameter()]
	[string] $TargetVersion ="16.0"

)
Begin{
    $defaultDisplaySet = 'GroupPolicy','Key', 'ValueName', 'Type', 'Value', 'Configuration'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

	$assemblies = ('System', 'mscorlib', 'System.IO');
	$sourceCode = @'
	// FROM: https://gallery.technet.microsoft.com/Read-or-modify-Registrypol-778fed6e
namespace TJX.PolFileEditor
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.IO;
    using Microsoft.Win32;
    public enum PolEntryType : uint
    {
        REG_NONE = 0,
        REG_SZ = 1,
        REG_EXPAND_SZ = 2,
        REG_BINARY = 3,
        REG_DWORD = 4,
        REG_DWORD_BIG_ENDIAN = 5,
        REG_MULTI_SZ = 7,
        REG_QWORD = 11,
    }
    public class PolEntry : IComparable<PolEntry>
    {
        private List<byte> byteList;
        public PolEntryType Type { get; set; }
        public string KeyName { get; set; }
        public string ValueName { get; set; }
        internal List<byte> DataBytes
        {
            get { return this.byteList; }
        }
        public uint DWORDValue
        {
            get
            {
                byte[] bytes = this.byteList.ToArray();
                switch (this.Type)
                {
                    case PolEntryType.REG_NONE:
                    case PolEntryType.REG_SZ:
                    case PolEntryType.REG_MULTI_SZ:
                    case PolEntryType.REG_EXPAND_SZ:
                        uint result;
                        if (UInt32.TryParse(this.StringValue, out result))
                        {
                            return result;
                        }
                        else
                        {
                            throw new InvalidCastException();
                        }
                    case PolEntryType.REG_DWORD:
                        if (bytes.Length != 4) { throw new InvalidOperationException(); }
                        if (BitConverter.IsLittleEndian == false) { Array.Reverse(bytes); }
                        return BitConverter.ToUInt32(bytes, 0);
                    case PolEntryType.REG_DWORD_BIG_ENDIAN:
                        if (bytes.Length != 4) { throw new InvalidOperationException(); }
                        if (BitConverter.IsLittleEndian == true) { Array.Reverse(bytes); }
                        return BitConverter.ToUInt32(bytes, 0);
                    case PolEntryType.REG_QWORD:
                        if (bytes.Length != 8) { throw new InvalidOperationException(); }
                        if (BitConverter.IsLittleEndian == false) { Array.Reverse(bytes); }
                        ulong lvalue = BitConverter.ToUInt64(bytes, 0);
                        if (lvalue > UInt32.MaxValue || lvalue < UInt32.MinValue)
                        {
                            throw new OverflowException("QWORD value '" + lvalue.ToString() + "' cannot fit into an UInt32 value.");
                        }
                        return (uint)lvalue;
                    case PolEntryType.REG_BINARY:
                        if (bytes.Length != 4) { throw new InvalidOperationException(); }
                        return BitConverter.ToUInt32(bytes, 0);
                    default:
                        throw new Exception("Reached default cast that should be unreachable in PolEntry.UIntValue");
                }
            }
            set
            {
                this.Type = PolEntryType.REG_DWORD;
                this.byteList.Clear();
                byte[] arrBytes = BitConverter.GetBytes(value);
                if (BitConverter.IsLittleEndian == false) { Array.Reverse(arrBytes); }
                this.byteList.AddRange(arrBytes);
            }
        }
        public ulong QWORDValue
        {
            get
            {
                byte[] bytes = this.byteList.ToArray();
                switch (this.Type)
                {
                    case PolEntryType.REG_NONE:
                    case PolEntryType.REG_SZ:
                    case PolEntryType.REG_MULTI_SZ:
                    case PolEntryType.REG_EXPAND_SZ:
                        ulong result;
                        if (UInt64.TryParse(this.StringValue, out result))
                        {
                            return result;
                        }
                        else
                        {
                            throw new InvalidCastException();
                        }
                    case PolEntryType.REG_DWORD:
                        if (bytes.Length != 4) { throw new InvalidOperationException(); }
                        if (BitConverter.IsLittleEndian == false) { Array.Reverse(bytes); }
                        return (ulong)BitConverter.ToUInt32(bytes, 0);
                    case PolEntryType.REG_DWORD_BIG_ENDIAN:
                        if (bytes.Length != 4) { throw new InvalidOperationException(); }
                        if (BitConverter.IsLittleEndian == true) { Array.Reverse(bytes); }
                        return (ulong)BitConverter.ToUInt32(bytes, 0);
                    case PolEntryType.REG_QWORD:
                        if (bytes.Length != 8) { throw new InvalidOperationException(); }
                        if (BitConverter.IsLittleEndian == false) { Array.Reverse(bytes); }
                        return BitConverter.ToUInt64(bytes, 0);
                    case PolEntryType.REG_BINARY:
                        if (bytes.Length != 8) { throw new InvalidOperationException(); }
                        return BitConverter.ToUInt64(bytes, 0);
                    default:
                        throw new Exception("Reached default cast that should be unreachable in PolEntry.ULongValue");
                }
            }
            set
            {
                this.Type = PolEntryType.REG_QWORD;
                this.byteList.Clear();
                byte[] arrBytes = BitConverter.GetBytes(value);
                if (BitConverter.IsLittleEndian == false) { Array.Reverse(arrBytes); }
                this.byteList.AddRange(arrBytes);
            }
        }
        public string StringValue
        {
            get
            {
                byte[] bytes = this.byteList.ToArray();
                StringBuilder sb = new StringBuilder(bytes.Length * 2);
                switch (this.Type)
                {
                    case PolEntryType.REG_NONE:
                        return "";
                    case PolEntryType.REG_MULTI_SZ:
                        string[] mstring = MultiStringValue;
                        for (int i = 0; i < mstring.Length; i++)
                        {
                            if (i > 0) { sb.Append("\\0"); }
                            sb.Append(mstring[i]);
                        }
                        return sb.ToString();
                    case PolEntryType.REG_DWORD:
                    case PolEntryType.REG_DWORD_BIG_ENDIAN:
                    case PolEntryType.REG_QWORD:
                        return this.QWORDValue.ToString();
                    case PolEntryType.REG_BINARY:
                        for (int i = 0; i < bytes.Length; i++)
                        {
                            sb.AppendFormat("{0:X2}", bytes[i]);
                        }
                        return sb.ToString();
                    case PolEntryType.REG_SZ:
                    case PolEntryType.REG_EXPAND_SZ:
                        return UnicodeEncoding.Unicode.GetString(bytes).Trim('\0');
                    default:
                        throw new Exception("Reached default cast that should be unreachable in PolEntry.StringValue");
                }
            }
            set
            {
                if (value == null) { value = String.Empty; }
                this.Type = PolEntryType.REG_SZ;
                this.byteList.Clear();
                this.byteList.AddRange(UnicodeEncoding.Unicode.GetBytes(value + "\0"));
            }
        }
        public string[] MultiStringValue
        {
            get
            {
                byte[] bytes = this.byteList.ToArray();
                switch (this.Type)
                {
                    case PolEntryType.REG_NONE:
                        throw new InvalidCastException("StringValue cannot be used on the REG_NONE type.");
                    case PolEntryType.REG_DWORD:
                    case PolEntryType.REG_DWORD_BIG_ENDIAN:
                    case PolEntryType.REG_QWORD:
                    case PolEntryType.REG_BINARY:
                    case PolEntryType.REG_SZ:
                    case PolEntryType.REG_EXPAND_SZ:
                        return new string[] { this.StringValue };
                    case PolEntryType.REG_MULTI_SZ:
                        List<string> list = new List<string>();
                        StringBuilder sb = new StringBuilder(256);
                        for (int i = 0; i < (bytes.Length - 1); i += 2)
                        {
                            char[] curChar = UnicodeEncoding.Unicode.GetChars(bytes, i, 2);
                            if (curChar[0] == '\0')
                            {
                                if (sb.Length == 0) { break; }
                                list.Add(sb.ToString());
                                sb.Length = 0;
                            }
                            else
                            {
                                sb.Append(curChar[0]);
                            }
                        }
                        return list.ToArray();
                    default:
                        throw new Exception("Reached default cast that should be unreachable in PolEntry.MultiStringValue");
                }
            }
            set
            {
                this.Type = PolEntryType.REG_MULTI_SZ;
                this.byteList.Clear();
                if (value != null)
                {
                    for (int i = 0; i < value.Length; i++)
                    {
                        if (i > 0) { this.byteList.AddRange(UnicodeEncoding.Unicode.GetBytes("\0")); }
                        if (value[i] != null)
                        {
                            this.byteList.AddRange(UnicodeEncoding.Unicode.GetBytes(value[i]));
                        }
                    }
                }
                this.byteList.AddRange(UnicodeEncoding.Unicode.GetBytes("\0\0"));
            }
        }
        public byte[] BinaryValue
        {
            get { return this.byteList.ToArray(); }
            set
            {
                this.Type = PolEntryType.REG_BINARY;
                this.byteList.Clear();
                if (value != null)
                {
                    this.byteList.AddRange(value);
                }
            }
        }
        public void SetDWORDBigEndianValue(uint value)
        {
            this.Type = PolEntryType.REG_DWORD_BIG_ENDIAN;
            this.byteList.Clear();
            byte[] arrBytes = BitConverter.GetBytes(value);
            if (BitConverter.IsLittleEndian == true) { Array.Reverse(arrBytes); }
            this.byteList.AddRange(arrBytes);
        }
        public void SetExpandStringValue(string value)
        {
            this.StringValue = value;
            this.Type = PolEntryType.REG_EXPAND_SZ;
        }
        public PolEntry()
        {
            this.byteList = new List<byte>();
            Type = PolEntryType.REG_NONE;
            KeyName = "";
            ValueName = "";
        }
        ~PolEntry()
        {
            this.byteList = null;
        }
        // IComparable<PolEntry>
        public int CompareTo(PolEntry other)
        {
            int result;
            result = String.Compare(this.KeyName, other.KeyName, StringComparison.OrdinalIgnoreCase);
            if (result != 0) { return result; }
            bool firstSpecial, secondSpecial;
            firstSpecial = this.ValueName.StartsWith("**", StringComparison.OrdinalIgnoreCase);
            secondSpecial = other.ValueName.StartsWith("**", StringComparison.OrdinalIgnoreCase);
            if (firstSpecial == true && secondSpecial == false) { return -1; }
            if (secondSpecial == true && firstSpecial == false) { return 1; }
            return String.Compare(this.ValueName, other.ValueName, StringComparison.OrdinalIgnoreCase);
        }
    }
    public class PolFile
    {
        private enum PolEntryParseState
        {
            Key,
            ValueName,
            Start
        }
        private static readonly uint PolHeader = 0x50526567;
        private static readonly uint PolVersion = 0x01000000;
        private Dictionary<string, PolEntry> entries;
        public List<PolEntry> Entries
        {
            get
            {
                List<PolEntry> pl = new List<PolEntry>(entries.Values);
                pl.Sort();
                return pl;
            }
        }
        public string FileName { get; set; }
        public PolFile()
        {
            this.FileName = "";
            this.entries = new Dictionary<string, PolEntry>(StringComparer.OrdinalIgnoreCase);
        }
        public void SetValue(PolEntry pe)
        {
            this.entries[pe.KeyName + "\\" + pe.ValueName] = pe;
        }
        public void SetStringValue(string key, string value, string data)
        {
            this.SetStringValue(key, value, data, false);
        }
        public void SetStringValue(string key, string value, string data, bool bExpand)
        {
            PolEntry pe = new PolEntry();
            pe.KeyName = key;
            pe.ValueName = value;
            if (bExpand)
            {
                pe.SetExpandStringValue(data);
            }
            else
            {
                pe.StringValue = data;
            }
            this.SetValue(pe);
        }
        public void SetDWORDValue(string key, string value, uint data)
        {
            this.SetDWORDValue(key, value, data, true);
        }
        public void SetDWORDValue(string key, string value, uint data, bool bLittleEndian)
        {
            PolEntry pe = new PolEntry();
            pe.KeyName = key;
            pe.ValueName = value;
            if (bLittleEndian)
            {
                pe.DWORDValue = data;
            }
            else
            {
                pe.SetDWORDBigEndianValue(data);
            }
            this.SetValue(pe);
        }
        public void SetQWORDValue(string key, string value, ulong data)
        {
            PolEntry pe = new PolEntry();
            pe.KeyName = key;
            pe.ValueName = value;
            pe.QWORDValue = data;
            this.SetValue(pe);
        }
        public void SetMultiStringValue(string key, string value, string[] data)
        {
            PolEntry pe = new PolEntry();
            pe.KeyName = key;
            pe.ValueName = value;
            pe.MultiStringValue = data;
            this.SetValue(pe);
        }
        public void SetBinaryValue(string key, string value, byte[] data)
        {
            PolEntry pe = new PolEntry();
            pe.KeyName = key;
            pe.ValueName = value;
            pe.BinaryValue = data;
            this.SetValue(pe);
        }
        public PolEntry GetValue(string key, string value)
        {
            PolEntry pe = null;
            this.entries.TryGetValue(key + "\\" + value, out pe);
            return pe;
        }
        public string GetStringValue(string key, string value)
        {
            PolEntry pe = this.GetValue(key, value);
            if (pe == null) { throw new ArgumentOutOfRangeException(); }
            return pe.StringValue;
        }
        public string[] GetMultiStringValue(string key, string value)
        {
            PolEntry pe = this.GetValue(key, value);
            if (pe == null) { throw new ArgumentOutOfRangeException(); }
            return pe.MultiStringValue;
        }
        public uint GetDWORDValue(string key, string value)
        {
            PolEntry pe = this.GetValue(key, value);
            if (pe == null) { throw new ArgumentOutOfRangeException(); }
            return pe.DWORDValue;
        }
        public ulong GetQWORDValue(string key, string value)
        {
            PolEntry pe = this.GetValue(key, value);
            if (pe == null) { throw new ArgumentOutOfRangeException(); }
            return pe.QWORDValue;
        }
        public byte[] GetBinaryValue(string key, string value)
        {
            PolEntry pe = this.GetValue(key, value);
            if (pe == null) { throw new ArgumentOutOfRangeException(); }
            return pe.BinaryValue;
        }
        public bool Contains(string key, string value)
        {
            return (this.GetValue(key, value) != null);
        }
        public bool Contains(string key, string value, PolEntryType type)
        {
            PolEntry pe = this.GetValue(key, value);
            return (pe != null && pe.Type == type);
        }
        public PolEntryType GetValueType(string key, string value)
        {
            PolEntry pe = this.GetValue(key, value);
            if (pe == null) { throw new ArgumentOutOfRangeException(); }
            return pe.Type;
        }
        public void DeleteValue(string key, string value)
        {
            if (this.entries.ContainsKey(key + "\\" + value) == true) this.entries.Remove(key + "\\" + value);
        }
        public void LoadFile()
        {
            this.LoadFile(null);
        }
        public void LoadFile(string file)
        {
            if (!string.IsNullOrEmpty(file)) { this.FileName = file; }
            byte[] bytes;
            int nBytes = 0;
            using (FileStream fs = new FileStream(this.FileName, FileMode.Open, FileAccess.Read))
            {
                // Read the source file into a byte array. 
                bytes = new byte[fs.Length];
                int nBytesToRead = (int)fs.Length;
                while (nBytesToRead > 0)
                {
                    // Read may return anything from 0 to nBytesToRead. 
                    int n = fs.Read(bytes, nBytes, nBytesToRead);
                    // Break when the end of the file is reached. 
                    if (n == 0) break;
                    nBytes += n;
                    nBytesToRead -= n;
                }
                fs.Close();
            }
            // registry.pol files are an 8-byte fixed header followed by some number of entries in the following format:
            // [KeyName;ValueName;<type>;<size>;<data>]
            // The brackets, semicolons, KeyName and ValueName are little-endian Unicode text.
            // type and size are 4-byte little-endian unsigned integers.  Size cannot be greater than 0xFFFF, even though it's
            // stored as a 32-bit number.  type will be one of the values REG_SZ, etc as defined in the Win32 API.
            // Data will be the number of bytes indicated by size.  The next 2 bytes afterward must be unicode "]".
            // 
            // All strings (KeyName, ValueName, and data when type is REG_SZ or REG_EXPAND_SZ) are terminated by a single
            // null character.
            //
            // Multi strings are strings separated by a single null character, with the whole list terminated by a double null.
            if (nBytes < 8) { throw new FileFormatException(); }
            int header = (bytes[0] << 24) | (bytes[1] << 16) | (bytes[2] << 8) | bytes[3];
            int version = (bytes[4] << 24) | (bytes[5] << 16) | (bytes[6] << 8) | bytes[7];
            if (header != PolFile.PolHeader || version != PolFile.PolVersion) { throw new FileFormatException(); }
            var parseState = PolEntryParseState.Start;
            int i = 8;
            var keyName = new StringBuilder(50);
            var valueName = new StringBuilder(50);
            uint type = 0;
            int size = 0;
            while (i < (nBytes - 1))
            {
                char[] curChar = UnicodeEncoding.Unicode.GetChars(bytes, i, 2);
                switch (parseState)
                {
                    case PolEntryParseState.Start:
                        if (curChar[0] != '[') { throw new FileFormatException(); }
                        i += 2;
                        parseState = PolEntryParseState.Key;
                        continue;
                    case PolEntryParseState.Key:
                        if (curChar[0] == '\0')
                        {
                            if (i > (nBytes - 4)) { throw new FileFormatException(); }
                            curChar = UnicodeEncoding.Unicode.GetChars(bytes, i + 2, 2);
                            if (curChar[0] != ';') { throw new FileFormatException(); }
                            // We've reached the end of the key name.  Switch to parsing value name.
                            i += 4;
                            parseState = PolEntryParseState.ValueName;
                        }
                        else
                        {
                            keyName.Append(curChar[0]);
                            i += 2;
                        }
                        continue;
                    case PolEntryParseState.ValueName:
                        if (curChar[0] == '\0')
                        {
                            if (i > (nBytes - 16)) { throw new FileFormatException(); }
                            curChar = UnicodeEncoding.Unicode.GetChars(bytes, i + 2, 2);
                            if (curChar[0] != ';') { throw new FileFormatException(); }
                            // We've reached the end of the value name.  Now read in the type and size fields, and the data bytes
                            type = (uint)(bytes[i + 7] << 24 | bytes[i + 6] << 16 | bytes[i + 5] << 8 | bytes[i + 4]);
                            if (Enum.IsDefined(typeof(PolEntryType), type) == false) { throw new FileFormatException(); }
                            curChar = UnicodeEncoding.Unicode.GetChars(bytes, i + 8, 2);
                            if (curChar[0] != ';') { throw new FileFormatException(); }
                            size = bytes[i + 13] << 24 | bytes[i + 12] << 16 | bytes[i + 11] << 8 | bytes[i + 10];
                            if ((size > 0xFFFF) || (size < 0)) { throw new FileFormatException(); }
                            curChar = UnicodeEncoding.Unicode.GetChars(bytes, i + 14, 2);
                            if (curChar[0] != ';') { throw new FileFormatException(); }
                            i += 16;
                            if (i > (nBytes - (size + 2))) { throw new FileFormatException(); }
                            curChar = UnicodeEncoding.Unicode.GetChars(bytes, i + size, 2);
                            if (curChar[0] != ']') { throw new FileFormatException(); }
                            PolEntry pe = new PolEntry();
                            pe.KeyName = keyName.ToString();
                            pe.ValueName = valueName.ToString();
                            pe.Type = (PolEntryType)type;
                            for (int j = 0; j < size; j++)
                            {
                                pe.DataBytes.Add(bytes[i + j]);
                            }
                            this.SetValue(pe);
                            i += size + 2;
                            keyName.Length = 0;
                            valueName.Length = 0;
                            parseState = PolEntryParseState.Start;
                        }
                        else
                        {
                            valueName.Append(curChar[0]);
                            i += 2;
                        }
                        continue;
                    default:
                        throw new Exception("Unreachable code");
                }
            }
        }
        public void SaveFile()
        {
            this.SaveFile(null);
        }
        public void SaveFile(string file)
        {
            if (!string.IsNullOrEmpty(file)) { this.FileName = file; }
            // Because we maintain the byte array for each PolEntry in memory, writing back to the file
            // is a simple operation, creating entries of the format:
            // [KeyName;ValueName;type;size;data] after the fixed 8-byte header.
            // The only things we must do are add null terminators to KeyName and ValueName, which are
            // represented by C# strings in memory, and make sure Size and Type are written in little-endian
            // byte order.
            using (FileStream fs = new FileStream(FileName, FileMode.Create, FileAccess.Write))
            {
                fs.Write(new byte[] { 0x50, 0x52, 0x65, 0x67, 0x01, 0x00, 0x00, 0x00 }, 0, 8);
                byte[] openBracket = UnicodeEncoding.Unicode.GetBytes("[");
                byte[] closeBracket = UnicodeEncoding.Unicode.GetBytes("]");
                byte[] semicolon = UnicodeEncoding.Unicode.GetBytes(";");
                byte[] nullChar = new byte[] { 0, 0 };
                byte[] bytes;
                foreach (PolEntry pe in this.Entries)
                {
                    fs.Write(openBracket, 0, 2);
                    bytes = UnicodeEncoding.Unicode.GetBytes(pe.KeyName);
                    fs.Write(bytes, 0, bytes.Length);
                    fs.Write(nullChar, 0, 2);
                    fs.Write(semicolon, 0, 2);
                    bytes = UnicodeEncoding.Unicode.GetBytes(pe.ValueName);
                    fs.Write(bytes, 0, bytes.Length);
                    fs.Write(nullChar, 0, 2);
                    fs.Write(semicolon, 0, 2);
                    bytes = BitConverter.GetBytes((uint)pe.Type);
                    if (BitConverter.IsLittleEndian == false) { Array.Reverse(bytes); }
                    fs.Write(bytes, 0, 4);
                    fs.Write(semicolon, 0, 2);
                    byte[] data = pe.DataBytes.ToArray();
                    bytes = BitConverter.GetBytes((uint)data.Length);
                    if (BitConverter.IsLittleEndian == false) { Array.Reverse(bytes); }
                    fs.Write(bytes, 0, 4);
                    fs.Write(semicolon, 0, 2);
                    fs.Write(data, 0, data.Length);
                    fs.Write(closeBracket, 0, 2);
                }
                fs.Close();
            }
        }
    }
    public class FileFormatException : Exception
    {
    }
}
'@;
}

Process
{
    Import-Module ServerManager

	#Get Policy file paths both machine and user
	$GPO = Get-GPO -Name $SourceGPOName;
    $TargetGPO = Get-GPO -Name $TargetGPOName -ErrorAction SilentlyContinue
    if (!($TargetGPO)) {
       $TargetGPO = New-GPO -Name $TargetGPOName
    }

    $TargetID = $TargetGPO.Id;
	$ID = $GPO.Id;
	$domain = [string]($GPO.DomainName);
   
    $sysvolDir = (Get-ItemProperty -Path hklm:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters).SysVol

	$Paths = [string]"$sysvolDir\$domain\Policies\{$ID}\User\Registry.pol", 
                     "$sysvolDir\$domain\Policies\{$ID}\Machine\Registry.pol";
    $commentPaths = ("$sysvolDir\$domain\Policies\{$ID}\User\comment.cmtx",
                     "$sysvolDir\$domain\Policies\{$TargetID}\User\comment.cmtx"),
                     ("$sysvolDir\$domain\Policies\{$ID}\Machine\comment.cmtx",
                     "$sysvolDir\$domain\Policies\{$TargetID}\Machine\comment.cmtx");
    $CentralPolicyDefinitionDirectory = "$sysvolDir\$domain\Policies\PolicyDefinitions"
	$PolicyDefinitionDirectory = "$env:windir\PolicyDefinitions"

	#Get all the admx files associated with the target version
    $admxPat = $TargetVersion.Substring(0,2);

    $admxFiles = dir -Path $CentralPolicyDefinitionDirectory -ErrorAction SilentlyContinue | ? Name -like "*$admxPat*";  

    if (!($admxFiles)) {
	    $admxFiles = dir -Path $PolicyDefinitionDirectory | ? Name -like "*$admxPat*";  
    }

	#Get all the policy definitions from the admx files
	$ExistingKeys = $null; 
    $j = 0;
	foreach($admx in $admxFiles){
        Write-Progress -Activity "Gathering Data from Administrative Templates" -status "Finding Registry Keys" -percentComplete ($j / $admxFiles.Count *100)
		[xml]$xml = Get-Content $admx.FullName;
		if($ExistingKeys -ne $null){
			$ExistingKeys += $xml.policyDefinitions.policies.policy | Select name, class, key, valueName;
		}else{
			$ExistingKeys = $xml.policyDefinitions.policies.policy | Select name, class, key, valueName;
		}
        $missingNames = $ExistingKeys | ? valueName -eq $null
        foreach($ek in $missingNames){
            $mnpolicy = $xml.policyDefinitions.policies.policy | ? name -EQ $ek.name | ? class -EQ $ek.class | ? key -EQ $ek.key
            foreach($mne in $mnpolicy.elements.ChildNodes){
                $ek.valueName = $mne.valueName;
                $ExistingKeys += $ek;
            }
        }
        $j = $j + 1;
	}

    try {
	      #Get Policy File reader object
	  Add-Type -TypeDefinition $sourceCode -ReferencedAssemblies $assemblies -ErrorAction SilentlyContinue;
    } catch { }

    $sourceCount=0;
    $targetCount=0;
    $results = new-object PSObject[] 0

	foreach($PolFilePath in $Paths)
 	{
        $ConfigType = "Machine"
        if ($PolFilePath -match '\\User\\') {
           $ConfigType = "User"
        }

        $fileExists = Test-Path -Path $PolFilePath
        if (!$fileExists) { continue }
		try 
		{ 
	        $pf = New-Object TJX.PolFileEditor.PolFile;
			$pf.LoadFile($PolFilePath) ;
		} 
		catch 
		{ 
			throw 
		} 

		$entries = [TJX.PolFileEditor.PolEntry[]] $pf.Entries.ToArray();

		#create regex patterns from inputs
		$SourceMatch = "office\\$SourceVersion";
		$SourceMatch = $SourceMatch -replace "\.", "\.";
		$TargetMatch = "office\\$TargetVersion";
		$TargetMatch = $TargetMatch -replace "\.", "\.";

		#Get entries corresponding to source and target locations
		[TJX.PolFileEditor.PolEntry[]]$entries_15 = @($entries | where {$_.KeyName -match $SourceMatch} )
		[TJX.PolFileEditor.PolEntry[]]$entries_16 = @($entries | where {$_.KeyName -match $TargetMatch} );


        if ($entries_15) { $sourceCount += $entries_15.Count; }
        if ($entries_16) { $targetCount += $entries_16.Count; }

        $i=0
        if($i -gt 0) {
           Write-Progress -Activity "Checking Group Policy Settings: $ConfigType" -status "Copying Settings..." -percentComplete ($i / $totalSettings*100)
        }
		#Find and copy each Policy but only if it doesn't already exist in the target location

        $totalSettings = $entries_15.Count;

		foreach($entry in $entries_15) 
		{
			if($entry.KeyName -match $SourceMatch)
			{
				$keyName = [string]$entry.KeyName;
				$newKeyName = $keyName.Replace($SourceVersion, $TargetVersion);

				[string]$key = "";
				[string]$compareClass = "";

				if($PolFilePath.Contains("Machine"))
				{
					$key = [string]"HKLM\$keyName";
					$compareClass = "Machine";
				}
				else
				{
					$key = [string]"HKCU\$keyName";
					$compareClass = "User";
				}       

				$valueName = $entry.ValueName;
        		$result = Get-GPRegistryValue -Name $SourceGpoName -Domain $domain -Key $key -ValueName $valueName;
				$hasValue = [System.Convert]::ToBoolean($result.HasValue);
				$type = [string]$result.Type.ToString();
				$newKey = $key.Replace($SourceVersion, $TargetVersion);

                #get value to compare if exists in admx files
                $compareKey = $newkey.Substring(5);

                $exists = $ExistingKeys | ? key -like $compareKey | ? valueName -eq $valueName;

                if($exists -ne $null){
				    $alreadySet = $false;
				    try
				    {
					    $existingKey = Get-GPRegistryValue -Name $SourceGpoName -Domain $domain -Key $newKey -ValueName $valueName -erroraction 'silentlycontinue';
                        $object =  New-Object PSObject -Property @{GroupPolicy = $TargetGpoName; Key = $newKey; ValueName = $valueName; Type = $type; Value = $result.Value; Configuration = $ConfigType }

                        $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers

					    if($existingKey.Value -eq $result.Value)
					    {
						    $alreadySet = $true;
                            Write-Progress -Activity "Checking Group Policy Settings: $ConfigType" -status "Setting Already Exists: $newKey\$valueName" -percentComplete ($i / $totalSettings*100)
					    }
				    }
				    catch
				    {
					    $alreadySet = $false;
				    }

				    if(!$alreadySet)
				    {
                        Write-Progress -Activity "Checking Group Policy Settings: $ConfigType" -status "Copying Setting: $newKey\$valueName" -percentComplete ($i / $totalSettings*100)

					    if($hasValue)
					    {
						   $SetValue = Set-GPRegistryValue -Name $TargetGpoName -Domain $domain -Key $newKey -ValueName $valueName -Type $type -Value $result.Value;
                           $object =  New-Object PSObject -Property @{GroupPolicy = $TargetGpoName; Key = $newKey; ValueName = $valueName; Type = $type; Value = $result.Value; Configuration = $ConfigType }
					    }
					    else
					    {
						   $SetValue = Set-GPRegistryValue -Name $TargetGpoName -Domain $domain -Key $newKey -ValueName $valueName -Type $type;
                           $object =  New-Object PSObject -Property @{GroupPolicy = $TargetGpoName; Key = $newKey; ValueName = $valueName; Type = $type; Configuration = $ConfigType }
					    }

                        $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                        $Results += $object
				    }


                }
			}    
          $i++
		}
        
	}

    $shortOldVersion = $SourceVersion.Split(".")[0];
    $shortNewVersion = $TargetVersion.Split(".")[0];

    foreach($commentFilePath in $commentPaths){
        if(Test-Path $commentFilePath[0]){
            [xml]$cmtx = Get-Content $commentFilePath[0];
            if(Test-Path $commentFilePath[1]){
                [xml]$newcmtx = Get-Content $commentFilePath[1];
            }else{
                $cmtx.Save($commentFilePath[1]);
                [xml]$newcmtx = Get-Content $commentFilePath[1];
                $newcmtx.policyComments.policyNamespaces.RemoveAll();
                $newcmtx.policyComments.comments.admTemplate.RemoveAll();
                $newcmtx.policyComments.resources.stringTable.RemoveAll();
            }

            $namespaces = $cmtx.policyComments.policyNamespaces.using | ? namespace -Like "*$shortOldVersion.Office.Microsoft.Policies.Windows";
            $newNamespaces = $newcmtx.policyComments.policyNamespaces.using | ? namespace -Like "*$shortNewVersion.Office.Microsoft.Policies.Windows";

            foreach($namespace in $namespaces){
                #get comments associated with each namespace
                $comments = $cmtx.policyComments.comments.admTemplate.comment | ? policyRef -Like "$($namespace.prefix):*"
                $resources = $cmtx.policyComments.resources.stringTable.string;
                
                #create new namespace to copy to
                $convertedNamespace = $namespace.namespace -replace "$shortOldVersion", "$shortNewVersion"
                $newNamespace = $newNamespaces | ? namespace -Like "$convertedNamespace";
                if($newNamespace -eq $null){
                    $maximum = $newcmtx.policyComments.policyNamespaces.using.prefix -replace "ns", "" | %{[int32]::Parse($_)} | measure -Maximum
                    $newNamespace = $newcmtx.CreateElement("using", $newcmtx.policyComments.NamespaceURI)
                    $newNamespace.SetAttribute("prefix", "ns$($maximum.Maximum + 1)") | Out-Null
                    $newNamespace.SetAttribute("namespace", $convertedNamespace) | Out-Null
                    $newcmtx.policyComments.GetElementsByTagName("policyNamespaces").AppendChild($newNamespace) | Out-Null
                }

                #copy all comments
                foreach($comment in $comments){
                    $nComment = $newcmtx.CreateElement("comment", $newcmtx.policyComments.NamespaceURI);
                    $nComment.RemoveAllAttributes();
                    #convert policy ref
                    $nPolicyRef = $comment.policyRef;
                    $nPolicyRef = $nPolicyRef -replace "$($namespace.prefix):", "$($newNamespace.prefix):"
                    $nComment.SetAttribute("policyRef", $nPolicyRef) | Out-Null

                    #convert comment text
                    $commentText = $comment.commentText -replace "$($namespace.prefix)_", "$($newNamespace.prefix)_"
                    $nComment.SetAttribute("commentText", $commentText) | Out-Null
                    $id = $comment.commentText -replace "\$\(resource\.", "";
                    $id = $id -replace "\)", "";
                    
                    $string = $resources | ? id -eq $id;
                    $string = $string.InnerText;

                    $id = $id -replace "$($namespace.prefix)_", "$($newNamespace.prefix)_"
                    $nResource = $newcmtx.CreateElement("string", $newcmtx.policyComments.NamespaceURI);
                    $nResource.RemoveAllAttributes();
                    $nResource.SetAttribute("id", $id) | Out-Null
                    $nResource.InnerText = $string;

                    #actually add the elements
                    $newcmtx.policyComments.comments.GetElementsByTagName("admTemplate").AppendChild($nComment) | Out-Null
                    $newcmtx.policyComments.resources.GetElementsByTagName("stringTable").AppendChild($nResource) | Out-Null
                }
            }
            $newcmtx.Save($commentFilePath[1]);
        }
    }
    
    Write-Host

    if ($sourceCount -eq 0 -and $targetCount -eq 0) {
        Write-Host "No Office settings are configured in the source Group Policy Object"
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "No Office settings are configured in the source Group Policy Object"
    } else {
        if ($Results.Count -eq 0) {
            Write-Host "All Office settings have already been copied"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "All Office settings have already been copied"
        }
    }

    $Results

}

}

function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

function Get-CurrentFunctionName {
    (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name;
}

Function WriteToLogFile() {
    param( 
      [Parameter(Mandatory=$true)]
      [string]$LNumber,
      [Parameter(Mandatory=$true)]
      [string]$FName,
      [Parameter(Mandatory=$true)]
      [string]$ActionError
    )
    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        #check if file exists, create if it doesn't
        $getCurrentDatePath = "C:\Windows\Temp\" + (Get-Date -Format u).Substring(0,10)+"OfficeAutoScriptLog.txt"
        if(Test-Path $getCurrentDatePath){#if exists, append  
             Add-Content $getCurrentDatePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $getCurrentDatePath $headerString
             Add-Content $getCurrentDatePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
    }
}

Copy-OfficeGPOSettings -SourceGPOName $SourceGPOName -TargetGPOName $TargetGPOName -SourceVersion $SourceVersion -TargetVersion $TargetVersion