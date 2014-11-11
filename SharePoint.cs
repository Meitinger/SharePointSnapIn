/* Copyright (C) 2014, Manuel Meitinger
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 2 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using SharePointSnapIn.Properties;

namespace SharePointSnapIn.SharePoint
{
    /// <summary>
    /// Represents the abstract root class for all SharePoint objects.
    /// </summary>
    public abstract class BaseObject
    {
        /// <summary>
        /// SharePoint Web Service namespace constant.
        /// </summary>
        public static readonly XNamespace SP = "http://schemas.microsoft.com/sharepoint/soap/";

        internal BaseObject(XElement xml)
        {
            // all SharePoint objects are derived from xml, so store the definition
            Xml = xml;
        }

        internal XElement Xml { get; private set; }

        private T Convert<T>(XName name, string value, bool isLookup, bool isRequired)
        {
            // remove the lookup prefix if necessary
            if (isLookup)
            {
                var sepIndex = value.IndexOf(";#");
                var dummy = default(int);
                if (sepIndex == -1 || !int.TryParse(value.Substring(0, sepIndex), out dummy))
                    throw new FormatException(string.Format(Resources.LookupPrefixInvalid, name));
                value = value.Substring(sepIndex + 2);
                if (value.Length == 0)
                {
                    // handle an empty value after prefix removal
                    if (isRequired)
                        throw new ArgumentException(string.Format(Resources.RequiredAttributeOrElementMissing, name), "name");
                    else
                        return default(T);
                }
            }

            // strip the nullable layer
            var type = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
            try
            {
                // handle enum types separately
                if (!type.IsEnum)
                {
                    // convert using ChangeType if possible
                    try { return (T)System.Convert.ChangeType(value, type, CultureInfo.InvariantCulture); }
                    catch (InvalidCastException)
                    {
                        // find a static parse method with a formatting provider
                        var parse = type.GetMethod("Parse", BindingFlags.Static | BindingFlags.Public, null, new Type[] { typeof(string), typeof(IFormatProvider) }, null);
                        if (parse == null)
                        {
                            // find a culture agnostic parse method
                            parse = type.GetMethod("Parse", BindingFlags.Static | BindingFlags.Public, null, new Type[] { typeof(string) }, null);
                            if (parse == null)
                            {
                                // find a constructor that takes a string
                                var constr = type.GetConstructor(new Type[] { typeof(string) });
                                if (constr == null)
                                    throw;
                                else
                                    return (T)constr.Invoke(new object[] { value });
                            }
                            else
                                return (T)parse.Invoke(null, new object[] { value });
                        }
                        else
                            return (T)parse.Invoke(null, new object[] { value, CultureInfo.InvariantCulture });
                    }
                }
                else
                    return (T)Enum.Parse(type, value, true);
            }
            catch (FormatException e) { throw new FormatException(string.Format(Resources.AttributeFormatInvalid, name, e.Message), e); }
            catch (OverflowException e) { throw new FormatException(string.Format(Resources.AttributeFormatInvalid, name, e.Message), e); }
        }

        /// <summary>
        /// Gets a mandatory attribute.
        /// </summary>
        /// <typeparam name="T">The attribute's type.</typeparam>
        /// <param name="name">The name of the attribute.</param>
        /// <param name="isLookup">Indicates whether the value includes a lookup prefix.</param>
        /// <returns>The attribute's value.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">No attribute with the given <paramref name="name"/> exists or <paramref name="isLookup"/> is set but no prefix is present.</exception>
        /// <exception cref="FormatException">The value's format is invalid.</exception>
        protected T Get<T>(XName name, bool isLookup = false)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            var attr = Xml.Attribute(name);
            if (attr == null || string.IsNullOrEmpty(attr.Value))
                throw new ArgumentException(string.Format(Resources.RequiredAttributeOrElementMissing, name), "name");
            return Convert<T>(name, attr.Value, isLookup, true);
        }

        /// <summary>
        /// Gets an optional attribute.
        /// </summary>
        /// <typeparam name="T">The attribute's type.</typeparam>
        /// <param name="name">The name of the attribute.</param>
        /// <param name="isLookup">Indicates whether the value includes a lookup prefix.</param>
        /// <returns>The attribute's value or <c>default(T)</c> if it's not specified.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException"><paramref name="isLookup"/> is set but no prefix is present.</exception>
        /// <exception cref="FormatException">The value's format is invalid.</exception>
        protected T GetOptional<T>(XName name, bool isLookup = false)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            var attr = Xml.Attribute(name);
            return attr == null || string.IsNullOrEmpty(attr.Value) ?
                default(T) :
                Convert<T>(name, attr.Value, isLookup, false);
        }

        /// <summary>
        /// Gets a mandatory element.
        /// </summary>
        /// <param name="name">The name of the element.</param>
        /// <returns>The element.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">No attribute with the given <paramref name="name"/> exists.</exception>
        protected XElement GetElement(XName name)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            var element = Xml.Element(name);
            if (element == null)
                throw new ArgumentException(string.Format(Resources.RequiredAttributeOrElementMissing, name), "name");
            return element;
        }
    }

    /// <summary>
    /// Specifies the type of a SharePoint list.
    /// </summary>
    public enum BaseType
    {
        /// <summary>
        /// A generic type of list template used for most lists.
        /// </summary>
        GenericList = 0,

        /// <summary>
        /// A document library.
        /// </summary>
        DocumentLibrary = 1,

        /// <summary>
        /// Unused.
        /// </summary>
        Unused = 2,

        /// <summary>
        /// A discussion board.  
        /// </summary>
        DiscussionBoard = 3,

        /// <summary>
        /// A survey list.
        /// </summary>
        Survey = 4,

        /// <summary>
        /// An issue-tracking list.
        /// </summary>
        Issue = 5,
    }

    /// <summary>
    /// Specifies the various list properties (see [MS-PRIMEPF]).
    /// </summary>
    [Flags]
    public enum ListFlags : ulong
    {
        OrderedList = 0x1,
        PublicList = 0x2,
        UndeleteableList = 0x4,
        DisableAttachments = 0x8,
        CatalogList = 0x10,
        MultipleMeetingDataList = 0x20,
        EnableAssignedToEmail = 0x40,
        VersioningEnabled = 0x80,
        HiddenList = 0x100,
        RequestAccessDenied = 0x200,
        ModeratedList = 0x400,
        AllowMultiVote = 0x800,
        UseForcedDisplay = 0x1000,
        DontSaveTemplate = 0x2000,
        RootWebOnly = 0x4000,
        MustSaveRootFiles = 0x8000,
        EmailInsertsEnabled = 0x10000,
        PrivateList = 0x20000,
        ForceCheckout = 0x40000,
        MinorVersionEnabled = 0x80000,
        MinorAuthor = 0x100000,
        MinorApprover = 0x200000,
        EnableContentTypes = 0x400000,
        FieldSchemaModified = 0x800000,
        EnableThumbnails = 0x1000000,
        AllowEveryoneViewItems = 0x2000000,
        WorkFlowAssociated = 0x4000000,
        DisableDeployWithDependentList = 0x8000000,
        DefaultItemOpen = 0x10000000,
        DisableFolders = 0x20000000,
        RestrictedTemplateList = 0x40000000,
        HasContentTypeOrder = 0x80000000,
        ForcedDefaultContentType = 0x100000000,
        CacheSchema = 0x200000000,
        IgnoreSealedAttribute = 0x400000000,
        NoCrawl = 0x800000000,
        AlwaysIncludeContent = 0x1000000000,
        DisallowContentTypes = 0x2000000000,
        SyndicationDisabled = 0x4000000000,
        IrmEnabled = 0x8000000000,
        IrmExpired = 0x10000000000,
        IrmRejected = 0x20000000000,
        EnablePeopleSelector = 0x80000000000,
        HasValidation = 0x100000000000,
        EnableSourceSelector = 0x200000000000,
        HasExternalDataSource = 0x400000000000,
        PreserveEmptyColumns = 0x800000000000,
        HasListScopedUserCustomActions = 0x1000000000000,
        ExcludeFromOfflineClient = 0x2000000000000,
        EnforceDataValidation = 0x4000000000000,
        DefaultItemOpenUseListSetting = 0x8000000000000,
        IsApplicationList = 0x10000000000000,
        DisableGridEditing = 0x20000000000000,
        BrowserFileHandlingStrict = 0x40000000000000,
        NavigateForFormSpaces = 0x80000000000000,
        StrictTypeCoercion = 0x100000000000000,
        EverEnabledDraft = 0x200000000000000,
        NeedUpdateSiteClientTag = 0x400000000000000,
        ReadOnlyUi = 0x800000000000000,
    }

    /// <summary>
    /// Represents a SharePoint list or document library.
    /// </summary>
    public sealed class List : BaseObject
    {
        internal List(SnapIn snapIn, XElement xml, IEnumerable<XElement> contentTypes)
            : base(xml)
        {
            // parse the xml element
            SnapIn = snapIn;
            Id = Get<Guid>("ID");
            BaseType = Get<BaseType>("BaseType");
            Flags = Get<ListFlags>("Flags");
            WebId = Get<Guid>("WebId");
            Version = Get<int>("Version");
            Path = Get<string>("RootFolder");
            WebPath = Get<string>("WebFullUrl");

            // create all enumerations
            Fields = Array.AsReadOnly(Field.Parse(this, GetElement(SP + "Fields").Elements(SP + "Field")).ToArray());
            ContentTypes = Array.AsReadOnly(contentTypes.Select(ct => new ContentType(this, ct)).ToArray());
        }

        /// <summary>
        /// Gets the snap-in that created the list.
        /// </summary>
        public SnapIn SnapIn { get; private set; }

        /// <summary>
        /// Gets the list's <see cref="Guid"/>.
        /// </summary>
        public Guid Id { get; private set; }

        /// <summary>
        /// Gets the type of the list.
        /// </summary>
        public BaseType BaseType { get; private set; }

        /// <summary>
        /// Gets the list's properties.
        /// </summary>
        /// <remarks>
        /// This is limit to properties with a value within an <see cref="int"/> range.
        /// </remarks>
        public ListFlags Flags { get; private set; }

        /// <summary>
        /// Gets the <see cref="Guid"/> of the SharePoint web that the list belongs to.
        /// </summary>
        public Guid WebId { get; private set; }

        /// <summary>
        /// Gets the list's version that is incremented every time its definition changes.
        /// </summary>
        public int Version { get; private set; }

        /// <summary>
        /// Gets the server-relative path to the list's root folder.
        /// </summary>
        public string Path { get; private set; }

        /// <summary>
        /// Gets the server-relative path to the SharePoint web that contains this list.
        /// </summary>
        public string WebPath { get; private set; }

        /// <summary>
        /// Gets an ordered list of all supported <see cref="Field"/>s defined in this list.
        /// </summary>
        public IList<Field> Fields { get; private set; }

        /// <summary>
        /// Gets an ordered list of all <see cref="ContentType"/>s.
        /// </summary>
        public IList<ContentType> ContentTypes { get; private set; }

        /// <summary>
        /// Gets an enumeration of all direct <see cref="Folder"/>s within the list.
        /// </summary>
        public IEnumerable<Folder> SubFolders { get { return SnapIn.GetListItems(WebId, Id, Folder.Query).Select(f => new Folder(this, null, f)); } }

        /// <summary>
        /// Gets an enumeration of all direct <see cref="ListItem"/>s within the list.
        /// </summary>
        public IEnumerable<ListItem> Items { get { return SnapIn.GetListItems(WebId, Id, ListItem.Query).Select(f => new ListItem(this, null, f)); } }
    }

    /// <summary>
    /// Represents a SharePoint content type.
    /// </summary>
    public sealed class ContentType : BaseObject
    {
        private class FieldStub : BaseObject
        {
            internal FieldStub(XElement xml)
                : base(xml)
            {
                Id = Get<Guid>("ID");
            }

            public Guid Id { get; private set; }
        }

        internal ContentType(List list, XElement xml)
            : base(xml)
        {
            // parse the xml element
            List = list;
            Id = Get<string>("ID");
            Name = Get<string>("Name");
            Version = Get<int>("Version");

            // create the fields
            Fields = Array.AsReadOnly(GetElement(SP + "Fields").Elements(SP + "Field").Select(f => new FieldStub(f).Id).SelectMany(id => list.Fields.Where(l => l.Id == id)).ToArray());
        }

        /// <summary>
        /// Gets the list this content type belongs to.
        /// </summary>
        public List List { get; private set; }

        /// <summary>
        /// Gets the identifier of the content type, starting with <c>0x</c>.
        /// </summary>
        public string Id { get; private set; }

        /// <summary>
        /// Gets the user-friendly name of the content type.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the content type's version that is incremented every time its definition changes.
        /// </summary>
        public int Version { get; private set; }

        /// <summary>
        /// Gets an ordered list of all the <see cref="Field"/>s in the <see cref="List"/> used by this content type.
        /// </summary>
        public IList<Field> Fields { get; private set; }
    }

    /// <summary>
    /// Specifies the type of a SharePoint <see cref="List"/> <see cref="Item"/>.
    /// </summary>
    public enum FileSystemObjectType
    {
        /// <summary>
        /// A document or list item.
        /// </summary>
        File = 0,

        /// <summary>
        /// A folder.
        /// </summary>
        Folder = 1,

        /// <summary>
        /// A SharePoint web.
        /// </summary>
        Web = 2,
    }

    /// <summary>
    /// Represents an item within a SharePoint <see cref="List"/>.
    /// </summary>
    public abstract class Item : BaseObject
    {
        private static readonly Regex UrlDescription = new Regex("[^,](,,)*?, ", RegexOptions.ExplicitCapture);

        private string GetUrlDescrption(string url)
        {
            // extracts the description from an url field value
            var match = UrlDescription.Match(url);
            return match.Success ? url.Substring(match.Index + match.Length) : url;
        }

        internal Item(List list, Folder rootFolder, XElement xml)
            : base(xml)
        {
            // parse the xml element
            List = list;
            RootFolder = rootFolder;
            Id = Get<int>("ows_ID");
            FileSystemObjectType = Get<FileSystemObjectType>("ows_FSObjType", true);
            Path = "/" + Get<string>("ows_FileRef", true);

            // find the content type
            var contentTypeId = Get<string>("ows_ContentTypeId");
            ContentType = List.ContentTypes.SingleOrDefault(ct => ct.Id == contentTypeId);

            // either use FileLeafRef, Title or URL description as name
            Name = Get<string>("ows_FileLeafRef", true);
            if (List.BaseType != BaseType.DocumentLibrary && FileSystemObjectType != FileSystemObjectType.Folder)
            {
                if (contentTypeId.StartsWith("0x0105", StringComparison.Ordinal))
                {
                    var url = GetOptional<string>("ows_URL");
                    if (url != null)
                        Name = GetUrlDescrption(Get<string>("ows_URL"));
                }
                else
                {
                    var title = GetOptional<string>("ows_Title");
                    if (title != null)
                        Name = title;
                }
            }
        }

        /// <summary>
        /// Gets the list this item belongs to.
        /// </summary>
        public List List { get; private set; }

        /// <summary>
        /// Gets the item's unique identifier within its list.
        /// </summary>
        public int Id { get; private set; }

        /// <summary>
        /// Gets the item's name, which is either its file name, title or link description.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the server-relative item path.
        /// </summary>
        public string Path { get; private set; }

        /// <summary>
        /// Gets the <see cref="Folder"/> this item belongs to or <c>null</c> if its parent is the <see cref="List"/> itself.
        /// </summary>
        public Folder RootFolder { get; private set; }

        /// <summary>
        /// Gets the type of this item.
        /// </summary>
        public FileSystemObjectType FileSystemObjectType { get; private set; }

        /// <summary>
        /// Gets the content type that this item belongs to.
        /// </summary>
        public ContentType ContentType { get; private set; }
    }

    /// <summary>
    /// Represents a folder within a SharePoint <see cref="List"/>.
    /// </summary>
    public sealed class Folder : Item
    {
        internal static readonly XElement Query =
            new XElement("Query",
                new XElement("Where",
                    new XElement("Eq",
                        new XElement("FieldRef",
                            new XAttribute("Name", "FSObjType")),
                        new XElement("Value", 1,
                            new XAttribute("Type", "Lookup")))),
                new XElement("OrderBy",
                    new XElement("FieldRef",
                        new XAttribute("Name", "FileRef"))));

        internal Folder(List list, Folder rootFolder, XElement xml)
            : base(list, rootFolder, xml)
        {
            var order = GetOptional<string>("ows_MetaInfo_vti_contenttypeorder");
            ContentTypeOrder = order == null ? list.ContentTypes : order.Split(',').SelectMany(id => list.ContentTypes.Where(ct => ct.Id == id)).ToList();
        }

        /// <summary>
        /// Gets the order in which the <see cref="List"/>'s <see cref="ContentType"/>s should be displayed.
        /// </summary>
        public IList<ContentType> ContentTypeOrder { get; private set; }

        /// <summary>
        /// Gets an enumeration of all direct <see cref="Folder"/>s within this folder.
        /// </summary>
        public IEnumerable<Folder> SubFolders { get { return List.SnapIn.GetListItems(List.WebId, List.Id, Query, Path, false).Select(f => new Folder(List, this, f)); } }

        /// <summary>
        /// Gets an enumeration of all direct <see cref="ListItem"/>s within this folder.
        /// </summary>
        public IEnumerable<ListItem> Items { get { return List.SnapIn.GetListItems(List.WebId, List.Id, ListItem.Query, Path, false).Select(f => new ListItem(List, this, f)); } }
    }

    /// <summary>
    /// Represents an item or file within a <see cref="List"/>.
    /// </summary>
    public sealed class ListItem : Item
    {
        internal static readonly XElement Query =
            new XElement("Query",
                new XElement("Where",
                    new XElement("Eq",
                        new XElement("FieldRef",
                            new XAttribute("Name", "FSObjType")),
                        new XElement("Value", 0,
                            new XAttribute("Type", "Lookup")))),
                new XElement("OrderBy",
                    new XElement("FieldRef",
                        new XAttribute("Name", "FileRef"))));

        internal ListItem(List list, Folder rootFolder, XElement xml) : base(list, rootFolder, xml) { }
    }

    /// <summary>
    /// Represents the base class for all fields within a SharePoint <see cref="List"/> or <see cref="ContentType"/>.
    /// </summary>
    public abstract class Field : BaseObject
    {
        private static bool Filter(XElement xml, XName attrName, bool expected)
        {
            // check if the given attribute matches the expected value
            var attr = xml.Attribute(attrName);
            return attr == null ? true : attr.Value != (expected ? "FALSE" : "TRUE");
        }

        internal static IEnumerable<Field> Parse(List list, IEnumerable<XElement> fields)
        {
            // create a typed field definition for each xml element
            foreach (var field in fields.Where(f => Filter(f, "ReadOnly", false) && Filter(f, "Hidden", false) && Filter(f, "ShowInNewForm", true)))
            {
                // try to get the type
                var typeAttr = field.Attribute("Type");
                if (typeAttr == null || string.IsNullOrEmpty(typeAttr.Value))
                    continue;
                switch (typeAttr.Value)
                {
                    case "Boolean":
                        yield return new BooleanField(list, field);
                        break;
                    case "Text":
                    case "Note":
                        yield return new TextField(list, field);
                        break;
                    case "Choice":
                        yield return new ChoiceField(list, field);
                        break;
                    case "MultiChoice":
                        yield return new MultipleChoiceField(list, field);
                        break;
                    case "Number":
                        yield return new NumberField(list, field);
                        break;
                    case "Currency":
                        yield return new CurrencyField(list, field);
                        break;
                    case "DateTime":
                        yield return new DateTimeField(list, field);
                        break;
                    case "Lookup":
                    case "LookupMulti":
                        yield return new LookupField(list, field);
                        break;
                    case "URL":
                        yield return new UrlField(list, field);
                        break;
                    case "User":
                    case "UserMulti":
                        yield return new UserField(list, field);
                        break;
                    case "Attachments":
                    case "File":
                        break;
                    default:
                        throw new NotSupportedException(string.Format(Resources.FieldTypeUnsupported, typeAttr.Value));
                }
            }
        }

        internal Field(List list, XElement xml)
            : base(xml)
        {
            // parse the common attributes
            List = list;
            Id = Get<Guid>("ID");
            InternalName = Get<string>("Name");
            DisplayName = Get<string>("DisplayName");
            var def = xml.Element(SP + "DefaultFormulaValue") ?? xml.Element(SP + "Default");
            DefaultValue = def == null ? null : def.Value;
            IsMandatory = GetOptional<bool>("Required");
        }

        /// <summary>
        /// Gets the list to which this field belongs to.
        /// </summary>
        public List List { get; private set; }

        /// <summary>
        /// Gets the <see cref="Guid"/> of the field.
        /// </summary>
        public Guid Id { get; private set; }

        /// <summary>
        /// Gets the internal name that is used for the field.
        /// </summary>
        public string DisplayName { get; private set; }

        /// <summary>
        /// Gets the display name for the field.
        /// </summary>
        public string InternalName { get; private set; }

        /// <summary>
        /// Gets the default value for a field.
        /// </summary>
        public string DefaultValue { get; private set; }

        /// <summary>
        /// Indicates whether the field requires a value.
        /// </summary>
        public bool IsMandatory { get; private set; }

        /// <summary>
        /// Gets suggestions based on the text the user has typed in so far.
        /// </summary>
        /// <param name="text">The input text.</param>
        /// <returns>All matching suggestions.</returns>
        public virtual IEnumerable<string> AutoComplete(string text)
        {
            return Enumerable.Empty<string>();
        }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.BaseField"/> that matches the field type.</returns>
        public abstract KMOAPICapture.BaseField CreateAutoStoreField();

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The AutoStore field.</param>
        /// <returns>A string value that fits this field's SharePoint schema or <c>null</c> if no value was specified.</returns>
        public abstract string Parse(KMOAPICapture.BaseField field);

        /// <summary>
        /// Retrieves all items of a dynamic list.
        /// </summary>
        /// <returns>All key-value pairs, where the key represents the item's AutoStore value and the value the text to be displayed.</returns>
        public virtual IEnumerable<KeyValuePair<string, string>> RetrieveItems()
        {
            return Enumerable.Empty<KeyValuePair<string, string>>();
        }
    }

    /// <summary>
    /// Represents a yes/no field.
    /// </summary>
    public sealed class BooleanField : Field
    {
        internal BooleanField(List list, XElement xml) : base(list, xml) { }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.CheckboxField"/>.</returns>
        public override KMOAPICapture.BaseField CreateAutoStoreField()
        {
            // create a checkbox field with culture specific true/false captions
            return new KMOAPICapture.CheckboxField()
            {
                TrueValue = Resources.BooleanTrueValue,
                FalseValue = Resources.BooleanFalseValue,
                BoolValue = DefaultValue == "1",
            };
        }

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The <see cref="KMOAPICapture.CheckboxField"/> returned by <see cref="BooleanField.CreateAutoStoreField"/>.</param>
        /// <returns><c>"1"</c> if the checkbox was checked, <c>"0"</c> otherwise.</returns>
        public override string Parse(KMOAPICapture.BaseField field)
        {
            // convert the culture specific value back to the SharePoint boolean value
            return (field as KMOAPICapture.CheckboxField).BoolValue ? "1" : "0";
        }
    }

    /// <summary>
    /// Represents the base class for (multiple) choice fields.
    /// </summary>
    public abstract class BaseChoiceField : Field
    {
        internal BaseChoiceField(List list, XElement xml)
            : base(list, xml)
        {
            HasFillInChoice = GetOptional<bool>("FillInChoice");
            Choices = Array.AsReadOnly(GetElement(SP + "CHOICES").Elements(SP + "CHOICE").Select(e => e.Value).ToArray());
        }

        /// <summary>
        /// Indicates whether the field allows custom user input.
        /// </summary>
        public bool HasFillInChoice { get; private set; }

        /// <summary>
        /// Gets an ordered list of all possible predefined choices.
        /// </summary>
        public IList<string> Choices { get; private set; }
    }

    /// <summary>
    /// Represents the base class for fields with numeric values.
    /// </summary>
    public abstract class BaseNumberField : Field
    {
        internal BaseNumberField(List list, XElement xml)
            : base(list, xml)
        {
            DecimalPlaces = GetOptional<int?>("Decimal");
            Minimum = GetOptional<double?>("Min");
            Maximum = GetOptional<double?>("Max");
        }

        /// <summary>
        /// Gets the number of decimal places to display or <c>null</c> if as many as necessary should be shown.
        /// </summary>
        public int? DecimalPlaces { get; private set; }

        /// <summary>
        /// Gets the minimum allowed value or <c>null</c> if there is no lower limit.
        /// </summary>
        public double? Minimum { get; private set; }

        /// <summary>
        /// Gets the maximum allowed value or <c>null</c> if there is no upper limit.
        /// </summary>
        public double? Maximum { get; private set; }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.TextField"/>.</returns>
        public override KMOAPICapture.BaseField CreateAutoStoreField()
        {
            // return a text field because numeric fields can't handle floating point input.
            return new KMOAPICapture.TextField()
            {
                Value = DefaultValue == null ? string.Empty : Format(double.Parse(DefaultValue, CultureInfo.InvariantCulture)),
                MaxChars = 255,
            };
        }

        /// <summary>
        /// Formats a numeric value into its string representation.
        /// </summary>
        /// <param name="value">The value to format.</param>
        /// <returns>The value's string as it would appear on SharePoint itself.</returns>
        protected abstract string Format(double value);

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The <see cref="KMOAPICapture.TextField"/> returned by <see cref="BaseNumberField.CreateAutoStoreField"/>.</param>
        /// <returns>A culture invariant string of the number or <c>null</c> if no value was specified.</returns>
        public override string Parse(KMOAPICapture.BaseField field)
        {
            // make sure there is an input
            var value = (field as KMOAPICapture.TextField).Value.Trim();
            if (value.Length == 0)
                return null;

            // parse the value and check if its within range
            double v = default(double);
            if (!TryParse(value, out v))
                throw new FormatException(string.Format(Resources.NumericFormatInvalid, Format(1234.56789)));
            if (Minimum.HasValue && v < Minimum.Value)
                throw new FormatException(string.Format(Resources.NumericMustBeAboveMinimum, Format(Minimum.Value)));
            if (Maximum.HasValue && v > Maximum.Value)
                throw new FormatException(string.Format(Resources.NumericMustBeBelowMaximum, Format(Maximum.Value)));

            // return the value as culture invariant string
            return v.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Tries to parse the user input into a numeric value.
        /// </summary>
        /// <param name="input">The culture and format specific string.</param>
        /// <param name="value">The parsed string's value.</param>
        /// <returns><c>true</c> if the string fits the format or <c>false</c> otherwise.</returns>
        protected abstract bool TryParse(string input, out double value);
    }

    /// <summary>
    /// Represents a field where one of multiple choices can be selected.
    /// </summary>
    public sealed class ChoiceField : BaseChoiceField
    {
        internal ChoiceField(List list, XElement xml) : base(list, xml) { }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.TextField"/> if <see cref="BaseChoiceField.HasFillInChoice"/> is <c>true</c>, a <see cref="KMOAPICapture.ListFieldEx"/> otherwise.</returns>
        public override KMOAPICapture.BaseField CreateAutoStoreField()
        {
            if (HasFillInChoice)
            {
                // create a textboc with static suggestions that override the user input
                return new KMOAPICapture.TextField()
                {
                    SuggestionListType = KMOAPICapture.TextSuggestionListType.List,
                    SuggestionList = new List<string>(Choices),
                    Value = DefaultValue ?? string.Empty,
                    MaxChars = 255,
                };
            }
            else
            {
                // create a single-select listbox and add the choices
                var result = new KMOAPICapture.ListFieldEx()
                {
                    AllowMultipleSelection = false,
                    RaiseFindEvent = false,
                };
                foreach (var choice in Choices)
                    result.Items.Add(new KMOAPICapture.ListItem(choice, choice, choice == DefaultValue));
                return result;
            }
        }

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The <see cref="KMOAPICapture.TextField"/> or <see cref="KMOAPICapture.ListFieldEx"/> returned by <see cref="ChoiceField.CreateAutoStoreField"/>.</param>
        /// <returns>The selected choice <c>null</c> if no value was entered or selected.</returns>
        public override string Parse(KMOAPICapture.BaseField field)
        {
            if (HasFillInChoice)
            {
                // trim the input and return it if it's not empty
                var value = (field as KMOAPICapture.TextField).Value.Trim();
                return value.Length == 0 ? null : value;
            }
            else
            {
                // return the selected choice if there is one
                var selected = (field as KMOAPICapture.ListFieldEx).Items.SingleOrDefault(e => e.Selected);
                return selected == null ? null : selected.Value;
            }
        }
    }

    /// <summary>
    /// Represents a field that requires a numeric currency value.
    /// </summary>
    public sealed class CurrencyField : BaseNumberField
    {
        internal CurrencyField(List list, XElement xml)
            : base(list, xml)
        {
            Culture = CultureInfo.GetCultureInfo(Get<int>("LCID"));
        }

        /// <summary>
        /// Gets the culture and currency that is used.
        /// </summary>
        public CultureInfo Culture { get; private set; }

        /// <summary>
        /// Formats a currency value into its string representation.
        /// </summary>
        /// <param name="value">The value to format.</param>
        /// <returns>The value's string as it would appear on SharePoint itself.</returns>
        protected override string Format(double value)
        {
            // format the value as culture specific currency string
            var format = "C";
            if (DecimalPlaces.HasValue)
                format += DecimalPlaces.Value.ToString(CultureInfo.InvariantCulture);
            return value.ToString(format, Culture);
        }

        /// <summary>
        /// Tries to parse the user input into a currency value.
        /// </summary>
        /// <param name="input">The currency string.</param>
        /// <param name="value">The parsed value.</param>
        /// <returns><c>true</c> if the string is a valid currency or <c>false</c> otherwise.</returns>
        protected override bool TryParse(string input, out double value)
        {
            // try to parse the user input as currency
            return double.TryParse(input, NumberStyles.Currency, Culture, out value);
        }
    }

    /// <summary>
    /// Indicates what the user can input in a <see cref="DateTimeField"/>.
    /// </summary>
    public enum DateTimeFormat
    {
        /// <summary>
        /// Specifies that only the date is included.
        /// </summary>
        DateOnly,

        /// <summary>
        /// Specifies that both date and time are included.
        /// </summary>
        DateTime,
    }

    /// <summary>
    /// Represents a date and time or date-only field.
    /// </summary>
    public sealed class DateTimeField : Field
    {
        private static readonly DateTimeOffset EmptyDateTime = new DateTime(1900, 1, 1, 0, 0, 0);
        private const string FormatString = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'";

        internal DateTimeField(List list, XElement xml)
            : base(list, xml)
        {
            Format = Get<DateTimeFormat>("Format");
        }

        /// <summary>
        /// Gets the <see cref="DateTimeFormat"/> being used.
        /// </summary>
        public DateTimeFormat Format { get; private set; }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.DateTimeField"/>.</returns>
        public override KMOAPICapture.BaseField CreateAutoStoreField()
        {
            // create a datetime field
            return new KMOAPICapture.DateTimeField()
            {
                DateTimeType = Format == DateTimeFormat.DateOnly ? KMOAPICapture.DateTimeFieldType.Date : KMOAPICapture.DateTimeFieldType.DateAndTime,
                DefaultToNow = false,
                DateTimeValue = DefaultValue != null ? DateTimeOffset.ParseExact(DefaultValue, FormatString, CultureInfo.InvariantCulture) : EmptyDateTime,
            };
        }

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The <see cref="KMOAPICapture.DateTimeField"/> returned by <see cref="DateTimeField.CreateAutoStoreField"/>.</param>
        /// <returns>The value in <c>yyyy-MM-ddTHH:mm:ss:Z</c> format or <c>null</c> if no date was selected.</returns>
        public override string Parse(KMOAPICapture.BaseField field)
        {
            // return the formatted value or null if "no" value was entered
            var value = (field as KMOAPICapture.DateTimeField).DateTimeValue;
            return value == EmptyDateTime ? null : value.ToString(FormatString, CultureInfo.InvariantCulture);
        }
    }

    /// <summary>
    /// Represents a field that allows the user to select value(s) from another SharePoint list.
    /// </summary>
    public sealed class LookupField : Field
    {
        private class Entry : BaseObject
        {
            internal Entry(string valueFieldName, XElement xml)
                : base(xml)
            {
                Id = Get<int>("ows_ID");
                Value = GetOptional<string>("ows_" + valueFieldName);
            }

            internal int Id { get; private set; }
            internal string Value { get; private set; }
        }

        internal LookupField(List list, XElement xml)
            : base(list, xml)
        {
            // parse the xml definition
            AllowMultipleValues = GetOptional<bool>("Mult");
            HasUnlimitedLength = GetOptional<bool>("UnlimitedLengthInDocumentLibrary");

            // make sure the list is either a valid guid or "Self"
            var sourceList = Get<string>("List");
            var guid = default(Guid);
            if (!Guid.TryParse(sourceList, out guid))
            {
                if (sourceList != "Self")
                    throw new NotSupportedException(string.Format(Resources.LookupListUnsupported, sourceList));
                SourceListId = list.Id;
            }
            else
                SourceListId = guid;

            // get the list's web and field
            SourceWebId = GetOptional<Guid?>("WebId") ?? list.WebId;
            SourceField = Get<string>("ShowField");
        }

        /// <summary>
        /// Indicates that multiple lookup entries can be selected.
        /// </summary>
        public bool AllowMultipleValues { get; private set; }

        /// <summary>
        /// Indicates that the amount of selected entries is not limited by their combined <see cref="SourceField"/> length.
        /// </summary>
        public bool HasUnlimitedLength { get; private set; }

        /// <summary>
        /// Gets the identifier of the SharePoint web that contains the referenced list.
        /// </summary>
        public Guid SourceWebId { get; private set; }

        /// <summary>
        /// Gets the referenced list's identifier.
        /// </summary>
        public Guid SourceListId { get; private set; }

        /// <summary>
        /// Gets the name of the field that should be displayed.
        /// </summary>
        public string SourceField { get; private set; }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.ListFieldEx"/>.</returns>
        public override KMOAPICapture.BaseField CreateAutoStoreField()
        {
            // create a dynamic list field
            return new KMOAPICapture.ListFieldEx()
            {
                AllowMultipleSelection = AllowMultipleValues,
                RaiseFindEvent = true,
            };
        }

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The <see cref="KMOAPICapture.ListFieldEx"/> returned by <see cref="LookupField.CreateAutoStoreField"/>.</param>
        /// <returns>The selected lookup id(s) or <c>null</c> if no entry was selected.</returns>
        public override string Parse(KMOAPICapture.BaseField field)
        {
            var items = (field as KMOAPICapture.ListFieldEx).Items;
            var value = (string)null;
            if (AllowMultipleValues)
            {
                // concat multiple values
                var selectedEntries = items.Where(e => e.Selected);
                if (!selectedEntries.Any())
                    return null;
                value = string.Join(";#", selectedEntries.Select(e => e.Value));
            }
            else
            {
                // return the single selected lookup value or null
                var selectedEntry = items.SingleOrDefault(e => e.Selected);
                if (selectedEntry == null)
                    return null;
                value = selectedEntry.Value;
            }

            // check the length of the value
            if (!HasUnlimitedLength && value.Length > 255)
                throw new FormatException(Resources.LookupValueTooLong);
            return value;
        }

        /// <summary>
        /// Retrieves all items in the source list.
        /// </summary>
        /// <returns>All key-value pairs, where the key represents the item's (escaped) lookup string the value the item's text to be displayed.</returns>
        public override IEnumerable<KeyValuePair<string, string>> RetrieveItems()
        {
            // create the query and fetch the entries
            var query =
                new XElement("Query",
                    new XElement("OrderBy",
                        new XElement("FieldRef", new XAttribute("Name", SourceField))));
            var entries = List.SnapIn.GetListItems(SourceWebId, SourceListId, query, null, true, "ID", SourceField).Select(e => new Entry(SourceField, e));

            // return a list item for all those with a value
            return entries.Where(e => e.Value != null).Select(e =>
            {
                // prepend the id and escape semi-colons if multiple values are allowed
                var key = e.Id.ToString(CultureInfo.InvariantCulture) + ";#" + (AllowMultipleValues ? e.Value.Replace(";", ";;") : e.Value);
                return new KeyValuePair<string, string>(key, e.Value);
            });
        }
    }

    /// <summary>
    /// Represents a multiple choice field.
    /// </summary>
    public sealed class MultipleChoiceField : BaseChoiceField
    {
        internal MultipleChoiceField(List list, XElement xml) : base(list, xml) { }

        /// <summary>
        /// Gets suggestions based on the text the user has typed in so far.
        /// </summary>
        /// <param name="text">The input text.</param>
        /// <returns>All matching choices.</returns>
        public override IEnumerable<string> AutoComplete(string text)
        {
            // get the last entry
            var lastEntryPos = text.LastIndexOf(";#");
            lastEntryPos = lastEntryPos == -1 ? 0 : lastEntryPos + 2;
            var lastEntry = text.Substring(lastEntryPos);
            text = text.Substring(0, lastEntryPos);

            // trim the last entry and find suitable suggestions
            var suggestions = Choices;
            lastEntry = lastEntry.Trim();
            if (lastEntry.Length > 0)
            {
                // find every item that contains the user's text
                suggestions = Choices.Where(c => c.IndexOf(lastEntry, StringComparison.CurrentCultureIgnoreCase) > -1).ToList();

                // add the text itself if it isn't a choice already
                if (!suggestions.Any(c => c.Equals(lastEntry, StringComparison.CurrentCultureIgnoreCase)))
                    suggestions.Insert(0, lastEntry);
            }

            // create a complete text for each suggestion
            return suggestions.Select(s => text + s + ";#");
        }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.TextField"/> if <see cref="BaseChoiceField.HasFillInChoice"/> is <c>true</c>, a <see cref="KMOAPICapture.ListFieldEx"/> otherwise.</returns>
        public override KMOAPICapture.BaseField CreateAutoStoreField()
        {
            if (HasFillInChoice)
            {
                // create a dynamic auto-complete textbox
                return new KMOAPICapture.TextField()
                {
                    IsSuggestionListDynamic = true,
                    SuggestionListType = KMOAPICapture.TextSuggestionListType.Search,
                    Value = (DefaultValue ?? string.Empty) + ";#",
                    MaxChars = 255,
                };
            }
            else
            {
                // create an multi-select listbox
                var result = new KMOAPICapture.ListFieldEx()
                {
                    AllowMultipleSelection = true,
                    RaiseFindEvent = false,
                };
                foreach (var choice in Choices)
                    result.Items.Add(new KMOAPICapture.ListItem(choice, choice, choice == DefaultValue));
                return result;
            }
        }

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The <see cref="KMOAPICapture.TextField"/> or <see cref="KMOAPICapture.ListFieldEx"/> returned by <see cref="MultipleChoiceField.CreateAutoStoreField"/>.</param>
        /// <returns>The concatenated values or <c>null</c> if no choice was entered or selected.</returns>
        public override string Parse(KMOAPICapture.BaseField field)
        {
            // get the selected values
            var entries = HasFillInChoice ?
                (field as KMOAPICapture.TextField).Value.Split(new string[] { ";#" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(s => s.Length > 0) :
                (field as KMOAPICapture.ListFieldEx).Items.Where(e => e.Selected).Select(e => e.Value);
            return entries.Any() ? ";#" + string.Join(";#", entries) + ";#" : null;
        }
    }

    /// <summary>
    /// Represents a general purpose text field.
    /// </summary>
    public sealed class TextField : Field
    {
        internal TextField(List list, XElement xml)
            : base(list, xml)
        {
            MaxLength = GetOptional<int?>("MaxLength") ?? 255;
        }

        /// <summary>
        /// Gets the maximum number of allowed input characters.
        /// </summary>
        public int MaxLength { get; private set; }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.TextField"/>.</returns>
        public override KMOAPICapture.BaseField CreateAutoStoreField()
        {
            // create a plain ole text field
            return new KMOAPICapture.TextField()
            {
                Value = DefaultValue ?? string.Empty,
                MaxChars = MaxLength,
            };
        }

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The <see cref="KMOAPICapture.TextField"/> returned by <see cref="TextField.CreateAutoStoreField"/>.</param>
        /// <returns>The trimmed value or <c>null</c> if no text was entered.</returns>
        public override string Parse(KMOAPICapture.BaseField field)
        {
            // make sure there is an input and return it
            var value = (field as KMOAPICapture.TextField).Value.Trim();
            return value.Length == 0 ? null : value;
        }
    }

    /// <summary>
    /// Represents a field with either an ordinary number or a percentage.
    /// </summary>
    public sealed class NumberField : BaseNumberField
    {
        internal NumberField(List list, XElement xml)
            : base(list, xml)
        {
            ShowAsPercentage = GetOptional<bool>("Percentage");
        }

        /// <summary>
        /// Indicates that the value should be multiplied by 100 and shown with a percent symbol.
        /// </summary>
        public bool ShowAsPercentage { get; private set; }

        /// <summary>
        /// Formats a number into its string representation.
        /// </summary>
        /// <param name="value">The value to format.</param>
        /// <returns>The value's string as it would appear on SharePoint itself.</returns>
        protected override string Format(double value)
        {
            // return the culture specific string
            var format = ShowAsPercentage ? "P" : "N";
            if (DecimalPlaces.HasValue)
                format += DecimalPlaces.Value.ToString(CultureInfo.InvariantCulture);
            return value.ToString(format, CultureInfo.CurrentCulture);
        }

        /// <summary>
        /// Tries to parse the user input into a number.
        /// </summary>
        /// <param name="input">The numeric string.</param>
        /// <param name="value">The parsed value.</param>
        /// <returns><c>true</c> if the string is a valid number or percentage, <c>false</c> otherwise.</returns>
        protected override bool TryParse(string input, out double value)
        {
            // just parse the culture specific value or...
            if (ShowAsPercentage)
            {
                // ...remove the % symbol first and devide the number by 100
                var i = input.IndexOf(CultureInfo.CurrentCulture.NumberFormat.PercentSymbol);
                if (i > -1)
                    input = input.Remove(i, CultureInfo.CurrentCulture.NumberFormat.PercentSymbol.Length);
                if (!double.TryParse(input, NumberStyles.Number, CultureInfo.CurrentCulture, out value))
                    return false;
                value /= 100;
                return true;
            }
            else
                return double.TryParse(input, NumberStyles.Number, CultureInfo.CurrentCulture, out value);
        }
    }

    /// <summary>
    /// Represents a URI field without a description.
    /// </summary>
    public sealed class UrlField : Field
    {
        internal UrlField(List list, XElement xml) : base(list, xml) { }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.TextField"/>.</returns>
        public override KMOAPICapture.BaseField CreateAutoStoreField()
        {
            // create a validating text field
            return new KMOAPICapture.TextField()
            {
                Value = DefaultValue ?? string.Empty,
                MaxChars = 255,
            };
        }

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The <see cref="KMOAPICapture.TextField"/> returned by <see cref="UrlField.CreateAutoStoreField"/>.</param>
        /// <returns>The escaped url including its description or <c>null</c> if no url was specified.</returns>
        public override string Parse(KMOAPICapture.BaseField field)
        {
            // make sure there is an input
            var value = (field as KMOAPICapture.TextField).Value.Trim();
            if (value.Length == 0)
                return null;

            // duplicate every colon and append the uri as description
            var uri = (Uri)null;
            if (!Uri.TryCreate(value, UriKind.Absolute, out uri))
                throw new FormatException(Resources.UrlFormatInvalid);
            return uri.AbsoluteUri.Replace(",", ",,") + ", " + uri.ToString();
        }
    }

    /// <summary>
    /// Indicates the various selectable principal types.
    /// </summary>
    public enum UserSelectionMode
    {
        /// <summary>
        /// Specifies that only inidividual users can be selected.
        /// </summary>
        PeopleOnly,

        /// <summary>
        /// Specifies that both individuals and groups can be selected.
        /// </summary>
        PeopleAndGroups,
    }

    /// <summary>
    /// Represents a lookup field especially for users and groups.
    /// </summary>
    public sealed class UserField : Field
    {
        internal UserField(List list, XElement xml)
            : base(list, xml)
        {
            AllowMultipleValues = GetOptional<bool>("Mult");
            if (GetOptional<int>("UserSelectionScope") != 0)
                throw new NotSupportedException(Resources.UserSelectionScopeUnsupported);
            SelectionMode = Get<UserSelectionMode>("UserSelectionMode");
        }

        /// <summary>
        /// Indicates whether multiple principals can be selected.
        /// </summary>
        public bool AllowMultipleValues { get; private set; }

        /// <summary>
        /// Indicates whether only users or also groups can be selected.
        /// </summary>
        public UserSelectionMode SelectionMode { get; private set; }

        /// <summary>
        /// Gets suggestions based on the text the user has typed in so far.
        /// </summary>
        /// <param name="text">The input text.</param>
        /// <returns>All matching user account names.</returns>
        public override IEnumerable<string> AutoComplete(string text)
        {
            // find the matching principals
            if (AllowMultipleValues)
            {
                // separate the last entry from the text
                var lastEntryPos = text.LastIndexOf(";");
                var lastEntry = text.Substring(lastEntryPos + 1);
                text = text.Substring(0, lastEntryPos + 1);

                // trim the last entry and find matching principals
                lastEntry = lastEntry.Trim();
                return List.SnapIn.SearchPrincipals(lastEntry, SelectionMode == UserSelectionMode.PeopleAndGroups).Select(p => text + p.AccountName + "; ");
            }
            else
            {
                // trim the text first
                text = text.Trim();
                return List.SnapIn.SearchPrincipals(text, SelectionMode == UserSelectionMode.PeopleAndGroups).Select(p => p.AccountName);
            }
        }

        /// <summary>
        /// Creates the corresponding AutoStore field.
        /// </summary>
        /// <returns>A <see cref="KMOAPICapture.TextField"/>.</returns>
        public override KMOAPICapture.BaseField CreateAutoStoreField()
        {
            // create a text field with dynamic auto-complete
            return new KMOAPICapture.TextField()
            {
                IsSuggestionListDynamic = true,
                SuggestionListType = KMOAPICapture.TextSuggestionListType.Search,
                Value = string.Empty,
                MaxChars = 255,
            };
        }

        /// <summary>
        /// Converts the AutoStore field's value to the corresponding SharePoint value.
        /// </summary>
        /// <param name="field">The <see cref="KMOAPICapture.TextField"/> returned by <see cref="UserField.CreateAutoStoreField"/>.</param>
        /// <returns>The principal(s) lookup value or <c>null</c> if no user or group was specified.</returns>
        public override string Parse(KMOAPICapture.BaseField field)
        {
            // get the trimmed names(s)
            var value = (field as KMOAPICapture.TextField).Value;
            var names = (AllowMultipleValues ? value.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries) : new string[] { value }).Select(s => s.Trim()).Where(s => s.Length > 0).ToArray();
            if (names.Length == 0)
                return null;

            // resolve and check the principals
            var principals = List.SnapIn.ResolvePrincipals(names, SelectionMode == UserSelectionMode.PeopleAndGroups);
            for (var i = 0; i < names.Length; i++)
            {
                if (!principals[i].IsResolved)
                    throw new FormatException(string.Format(Resources.UserNotResolved, names[i]));
                if (principals[i].MoreMatches != null && principals[i].MoreMatches.Length > 0)
                    throw new FormatException(string.Format(Resources.MultipleUserMatchesFound, names[i]));
            }

            // return the ids
            return string.Join(";#", principals.Select(p => p.UserInfoID).Distinct().Select(id => id.ToString(CultureInfo.InvariantCulture) + ";#"));
        }
    }
}
