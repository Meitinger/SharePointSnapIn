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
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.Xml.Linq;
using SharePointSnapIn.Properties;

namespace SharePointSnapIn
{
    /// <summary>
    /// Indicates how new documents are added to SharePoint.
    /// </summary>
    public enum SnapInAppendMode
    {
        /// <summary>
        /// Don't show the append mode checkbox and never append.
        /// </summary>
        Never,

        /// <summary>
        /// Don't show the append mode checkbox and always append.
        /// </summary>
        Always,

        /// <summary>
        /// Show the append mode checkbox and let the user decide.
        /// </summary>
        AskUser,
    }

    /// <summary>
    /// Indicates the operation to be performed on a SharePoint item.
    /// </summary>
    public enum SnapInListItemCommand
    {
        /// <summary>
        /// Create a new item.
        /// </summary>
        New,

        /// <summary>
        /// Update an existing item.
        /// </summary>
        Update,

        /// <summary>
        /// Delete an existing item.
        /// </summary>
        Delete,
    }

    /// <summary>
    /// Represents the snap-in configuration that can be changed depending on the current user or other parameters.
    /// </summary>
    public class SnapInSettings
    {
        private static readonly Guid ContentTypeId = new Guid("{c042a256-787d-4a6f-8a8a-cf6ab767f12d}");
        internal static readonly char[] InvalidFileNameChars;

        static SnapInSettings()
        {
            // get the invalid file name chars and make sure they include '/'
            InvalidFileNameChars = Path.GetInvalidFileNameChars();
            if (Array.IndexOf(InvalidFileNameChars, '/') == -1)
            {
                var len = InvalidFileNameChars.Length;
                Array.Resize(ref InvalidFileNameChars, len + 1);
                InvalidFileNameChars[len] = '/';
            }
        }

        private readonly Dictionary<string, bool> allowedContentTypes = new Dictionary<string, bool>();
        private readonly Dictionary<string, bool> deniedContentTypes = new Dictionary<string, bool>();
        private readonly HashSet<Guid> ignoredFieldIds = new HashSet<Guid>();
        private readonly HashSet<string> ignoredFieldInternalNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly SharePoint.List list;
        private bool isLocked = false;
        private SnapInAppendMode appendMode;
        private SharePoint.Folder rootFolder;
        private bool allowSubFolders;
        private string defaultFileName;
        private string fileExtension;

        internal SnapInSettings(SharePoint.List list)
        {
            this.list = list;
            AppendMode = list.BaseType == SharePoint.BaseType.DocumentLibrary ? SnapInAppendMode.Never : SnapInAppendMode.AskUser;
            RootFolderPath = "/";
            AllowSubFolders = true;
            DenyContentType("0x0120", true);
            DefaultFileName = string.Empty;
            FileExtension = "pdf";
            IgnoreField(ContentTypeId);
        }

        private void CheckLocked()
        {
            // make sure the object isn't locked
            if (isLocked)
                throw new InvalidOperationException();
        }

        internal void Lock()
        {
            // lock the object
            CheckLocked();
            isLocked = true;
        }

        /// <summary>
        /// Sets the way how scanned documents should be uploaded to SharePoint.
        /// </summary>
        /// <exception cref="InvalidEnumArgumentException">The value is not defined.</exception>
        /// <exception cref="ArgumentException">The value is not supported by the list type.</exception>
        public SnapInAppendMode AppendMode
        {
            internal get { return appendMode; }
            set
            {
                // check and set the value
                if (!Enum.IsDefined(typeof(SnapInAppendMode), value))
                    throw new InvalidEnumArgumentException("AppendMode", (int)value, typeof(SnapInAppendMode));
                if (list.BaseType == SharePoint.BaseType.DocumentLibrary && value != SnapInAppendMode.Never)
                    throw new ArgumentException(Resources.AppendLibraryUnsupported);
                CheckLocked();
                appendMode = value;
            }
        }

        /// <summary>
        /// Sets the default folder path relative to the SharePoint list or document library.
        /// </summary>
        /// <exception cref="ArgumentNullException">The value is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">A folder with the given name doesn't exist.</exception>
        public string RootFolderPath
        {
            internal get { return rootFolder == null ? list.Path : rootFolder.Path; }
            set
            {
                // normalize the path
                if (value == null)
                    throw new ArgumentNullException("RootFolderPath");
                var pieces = new List<string>(value.Split(new char[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, '/' }, StringSplitOptions.RemoveEmptyEntries));
                var position = 0;
                while (position < pieces.Count)
                {
                    switch (pieces[position])
                    {
                        case ".":
                            pieces.RemoveAt(position);
                            break;
                        case "..":
                            pieces.RemoveAt(position);
                            if (position > 0)
                                pieces.RemoveAt(--position);
                            break;
                        default:
                            position++;
                            break;
                    }
                }
                value = string.Join("/", pieces);

                // find the root folder
                if (value.Length > 0)
                {
                    var path = list.Path + "/" + value;
                    var query = new XElement("Query",
                        new XElement("Where",
                            new XElement("And",
                                new XElement("Eq",
                                    new XElement("FieldRef", new XAttribute("Name", "FSObjType")),
                                    new XElement("Value", new XAttribute("Type", "Lookup"), 1)),
                                new XElement("Eq",
                                    new XElement("FieldRef", new XAttribute("Name", "FileRef")),
                                    new XElement("Value", new XAttribute("Type", "Lookup"), path.Substring(1))))));
                    var folderXml = list.SnapIn.GetListItems(list.WebId, list.Id, query).SingleOrDefault();
                    if (folderXml == null)
                        throw new ArgumentException(string.Format(Resources.RootFolderNotFound, path), "RootFolderPath");
                    CheckLocked();
                    rootFolder = new SharePoint.Folder(list, null, folderXml);
                }
                else
                {
                    CheckLocked();
                    rootFolder = null;
                }
            }
        }

        internal SharePoint.Folder RootFolder { get { return rootFolder; } }

        /// <summary>
        /// Sets whether the user may change the path to a sub folder.
        /// </summary>
        public bool AllowSubFolders
        {
            internal get { return allowSubFolders; }
            set
            {
                CheckLocked();
                allowSubFolders = value;
            }
        }

        /// <summary>
        /// Allows content type(s) with the given id.
        /// </summary>
        /// <param name="id">The SharePoint identifier starting with <c>0x</c>.</param>
        /// <param name="inherited">Specifies that all derived content types are allowed as well.</param>
        /// <exception cref="ArgumentNullException"><paramref name="id"/> is <c>null</c>.</exception>
        public void AllowContentType(string id, bool inherited = true)
        {
            if (id == null)
                throw new ArgumentNullException("id");
            CheckLocked();
            if (inherited || !allowedContentTypes.ContainsKey(id))
                allowedContentTypes[id] = inherited;
        }

        /// <summary>
        /// Denies content type(s) with the given id.
        /// </summary>
        /// <param name="id">The SharePoint identifier starting with <c>0x</c>.</param>
        /// <param name="inherited">Specifies that all derived content types are denied as well.</param>
        /// <exception cref="ArgumentNullException"><paramref name="id"/> is <c>null</c>.</exception>
        public void DenyContentType(string id, bool inherited = true)
        {
            if (id == null)
                throw new ArgumentNullException("id");
            CheckLocked();
            if (inherited || !deniedContentTypes.ContainsKey(id))
                deniedContentTypes[id] = inherited;
        }

        internal IEnumerable<SharePoint.ContentType> FilterContentTypes(IEnumerable<SharePoint.ContentType> cts)
        {
            // check the denied and allowed filters
            foreach (var ct in cts)
            {
                var checker = new Func<KeyValuePair<string, bool>, bool>(f => f.Value ? ct.Id.StartsWith(f.Key, StringComparison.Ordinal) : ct.Id == f.Key);
                if (deniedContentTypes.Any(checker))
                    continue;
                if (allowedContentTypes.Count > 0 && !allowedContentTypes.Any(checker))
                    continue;
                yield return ct;
            }
        }

        /// <summary>
        /// Ignores a field by its id.
        /// </summary>
        /// <param name="id">The id of the field to ignore.</param>
        /// <exception cref="ArgumentNullException"><paramref name="id"/> is <see cref="Guid.Empty"/>.</exception>
        public void IgnoreField(Guid id)
        {
            if (id == Guid.Empty)
                throw new ArgumentNullException("id");
            CheckLocked();
            ignoredFieldIds.Add(id);
        }

        /// <summary>
        /// Ignores a field by its internal name.
        /// </summary>
        /// <param name="internalName">The internal name of the field to ignore.</param>
        /// <exception cref="ArgumentNullException"><paramref name="internalName"/> is <c>null</c>.</exception>
        public void IgnoreField(string internalName)
        {
            if (internalName == null)
                throw new ArgumentNullException("internalName");
            CheckLocked();
            ignoredFieldInternalNames.Add(internalName);
        }

        internal IEnumerable<SharePoint.Field> FilterFields(IEnumerable<SharePoint.Field> fields)
        {
            // make sure the field isn't filtered
            return fields.Where(f => !ignoredFieldIds.Contains(f.Id) && !ignoredFieldInternalNames.Contains(f.InternalName));
        }

        /// <summary>
        /// Sets the default file name (without extension).
        /// </summary>
        /// <exception cref="ArgumentNullException">The value is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">The value contains invalid characters.</exception>
        public string DefaultFileName
        {
            internal get { return defaultFileName; }
            set
            {
                // make sure the value only contains valid file name characters
                if (value == null)
                    throw new ArgumentNullException("DefaultFileName");
                if (value.IndexOfAny(InvalidFileNameChars) > -1)
                    throw new ArgumentException(string.Format(Resources.InvalidCharsInFileName, string.Join(" ", InvalidFileNameChars)), "DefaultFileName");
                CheckLocked();
                defaultFileName = value;
            }
        }

        /// <summary>
        /// Sets the extension that is appended to the file name.
        /// </summary>
        /// <exception cref="ArgumentNullException">The value is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">The value contains invalid characters.</exception>
        public string FileExtension
        {
            internal get { return fileExtension; }
            set
            {
                // make sure the value only contains valid file name characters and remove the leading '.'
                if (value == null)
                    throw new ArgumentNullException("FileExtension");
                if (value.IndexOfAny(InvalidFileNameChars) > -1)
                    throw new ArgumentException(string.Format(Resources.InvalidCharsInFileName, string.Join(" ", InvalidFileNameChars)), "FileExtension");
                if (value.Length > 0 && value[0] == '.')
                    value = value.Substring(1);
                CheckLocked();
                fileExtension = value;
            }
        }
    }

    /// <summary>
    /// Base class for all dynamic form Snap-Ins.
    /// </summary>
    public abstract class SnapIn : KMOAPICapture.ISnapInModule
    {
        private static readonly byte[] emptyFile = new byte[1] { 0x00 };
        private static readonly Dictionary<Guid, SharePoint.List> listCache = new Dictionary<Guid, SharePoint.List>();
        private const string ContextFieldId = "SP.CONTEXT";
        private const string AppendModeFieldId = "SP.APPENDMODE";
        private const string FolderFieldId = "SP.FOLDER";
        private const string ListItemFieldId = "SP.LISTITEM";
        private const string ContentTypeFieldId = "SP.CONTENTTYPE";
        private const string FileNameFieldId = "SP.FILENAME";
        private static readonly HashSet<string> BuiltInFieldIds = new HashSet<string>()
        {
            ContextFieldId, 
            AppendModeFieldId,
            FolderFieldId,
            ListItemFieldId,
            ContentTypeFieldId,
            FileNameFieldId,
        };

        private class ListStub : SharePoint.BaseObject
        {
            internal ListStub(XElement xml)
                : base(xml)
            {
                Version = Get<int>("Version");
            }

            public int Version { get; private set; }
        }

        private class ContentTypeStub : SharePoint.BaseObject
        {
            internal ContentTypeStub(XElement xml)
                : base(xml)
            {
                Id = Get<string>("ID");
                Version = Get<int>("Version");
            }

            public string Id { get; private set; }
            public int Version { get; private set; }
        }

        /// <summary>
        /// Represents the currently selected values.
        /// </summary>
        public sealed class Context
        {
            private class ContextField : KMOAPICapture.LabelField
            {
                internal ContextField(Context context)
                {
                    Context = context;
                    LabelText = string.Empty;
                }

                private ContextField(ContextField baseField)
                    : base(baseField)
                {
                    // clone the context
                    Context = new Context()
                    {
                        UserState = Context.UserState,
                        List = Context.List,
                        Settings = Context.Settings,
                        SelectedAppending = Context.SelectedAppending,
                        SelectedSubFolder = Context.SelectedSubFolder,
                        SelectedListItem = Context.SelectedListItem,
                        SelectedContentType = Context.SelectedContentType,
                        SelectedFileName = Context.SelectedFileName,
                        firstUpdate = Context.firstUpdate,
                        fieldMap = new Dictionary<string, SharePoint.Field>(Context.fieldMap),
                        cachedFolders = new Dictionary<string, SharePoint.Folder>(Context.cachedFolders),
                        cachedListItems = new Dictionary<string, SharePoint.ListItem>(Context.cachedListItems),
                    };
                    LabelText = baseField.LabelText;
                }

                internal Context Context { get; private set; }

                public override object Clone()
                {
                    return new ContextField(this);
                }
            }

            private bool firstUpdate = true;
            private Dictionary<string, SharePoint.Field> fieldMap = new Dictionary<string, SharePoint.Field>();
            private Dictionary<string, SharePoint.Folder> cachedFolders = new Dictionary<string, SharePoint.Folder>();
            private Dictionary<string, SharePoint.ListItem> cachedListItems = new Dictionary<string, SharePoint.ListItem>();

            internal static Context Get(KMOAPICapture.AsForm form)
            {
                // retrieves the field and set the form if it's been cloned
                var context = ((ContextField)form.Fields[ContextFieldId]).Context;
                if (context.Form == null)
                    context.Form = form;
                return context;
            }

            internal static void Create(KMOAPICapture.AsForm form, object userState, SharePoint.List list, SnapInSettings settings)
            {
                // lock the settings and create the builtin fields
                settings.Lock();
                var index = 0;

                // append field
                form.Fields.Insert(index++, new KMOAPICapture.CheckboxField()
                {
                    Name = AppendModeFieldId,
                    IsRequired = true,
                    IsHidden = settings.AppendMode != SnapInAppendMode.AskUser,
                    RaiseChangeEvent = true,
                    Display = Resources.AppendModeField,
                    FalseValue = Resources.BooleanFalseValue,
                    TrueValue = Resources.BooleanTrueValue,
                    BoolValue = settings.AppendMode == SnapInAppendMode.Always,
                });

                // folder
                var foldersList = new KMOAPICapture.ListField()
                {
                    Name = FolderFieldId,
                    IsRequired = true,
                    IsHidden = settings.AllowSubFolders == false,
                    RaiseChangeEvent = true,
                    Display = Resources.FolderField,
                    RaiseFindEvent = settings.AllowSubFolders,
                    AllowMultipleSelection = false,
                };
                form.Fields.Insert(index++, foldersList);

                // list item
                form.Fields.Insert(index++, new KMOAPICapture.ListFieldEx()
                {
                    Name = ListItemFieldId,
                    IsRequired = false,
                    IsHidden = true,
                    RaiseChangeEvent = true,
                    Display = Resources.ListItemField,
                    RaiseFindEvent = false,
                    AllowMultipleSelection = false,
                });

                // content type
                form.Fields.Insert(index++, new KMOAPICapture.ListField()
                {
                    Name = ContentTypeFieldId,
                    IsRequired = false,
                    IsHidden = true,
                    RaiseChangeEvent = true,
                    Display = Resources.ContentTypeField,
                    RaiseFindEvent = false,
                    AllowMultipleSelection = false,
                });

                // file name
                form.Fields.Insert(index++, new KMOAPICapture.TextField()
                {
                    Name = FileNameFieldId,
                    IsRequired = true,
                    IsHidden = false,
                    RaiseChangeEvent = true,
                    Display = Resources.FileNameField,
                    MaxChars = 255,
                    Value = settings.DefaultFileName,
                });

                // the context
                var context = new Context()
                {
                    UserState = userState,
                    List = list,
                    Settings = settings,
                };
                form.Fields.Insert(index++, new ContextField(context)
                {
                    Name = ContextFieldId,
                    IsRequired = false,
                    IsHidden = true,
                });

                // add the root folder
                context.cachedFolders.Add(settings.RootFolderPath, null);
                foldersList.Items.Add(new KMOAPICapture.ListItem(settings.RootFolderPath, "/", true));
            }

            private Context() { }

            internal SnapInSettings Settings { get; private set; }

            /// <summary>
            /// Gets the user state return by <see cref="SnapIn.Initialize"/>.
            /// </summary>
            public object UserState { get; private set; }

            /// <summary>
            /// Gets the underlying form.
            /// </summary>
            public KMOAPICapture.AsForm Form { get; private set; }

            /// <summary>
            /// Gets the underlying list.
            /// </summary>
            public SharePoint.List List { get; private set; }

            /// <summary>
            /// Indicates whether the user selected append mode.
            /// </summary>
            public bool SelectedAppending { get; private set; }

            /// <summary>
            /// Gets the selected sub-folder or <c>null</c> if the root folder is selected.
            /// </summary>
            public SharePoint.Folder SelectedSubFolder { get; private set; }

            /// <summary>
            /// Gets the selected list item or <c>null</c> if nothing is selected.
            /// </summary>
            public SharePoint.ListItem SelectedListItem { get; private set; }

            /// <summary>
            /// Gets the selected content type of <c>null</c> if no suitable content type is available.
            /// </summary>
            public SharePoint.ContentType SelectedContentType { get; private set; }

            /// <summary>
            /// Gets the selected file name without extension.
            /// </summary>
            public string SelectedFileName { get; private set; }


            internal SharePoint.Field Field(KMOAPICapture.BaseField field)
            {
                // return the matching SharePoint field
                var sharePointField = (SharePoint.Field)null;
                fieldMap.TryGetValue(field.Name, out sharePointField);
                return sharePointField;
            }

            internal IEnumerable<KMOAPICapture.ListItem> RetrieveFolders()
            {
                // fetch the folders
                var subFoldersXml = List.SnapIn.GetListItems(List.WebId, List.Id, SharePoint.Folder.Query, Settings.RootFolderPath, true);
                var subFolders = subFoldersXml.Select(f => new SharePoint.Folder(List, null, f));
                foreach (var subFolder in subFolders)
                {
                    cachedFolders.Add(subFolder.Path, subFolder);
                    yield return new KMOAPICapture.ListItem(subFolder.Path, subFolder.Path.Substring(Settings.RootFolderPath.Length));
                }
            }

            internal IEnumerable<KMOAPICapture.ListItem> RetrieveItems()
            {
                // fetch the items under the currently selected folder
                var listItemsXml = List.SnapIn.GetListItems(List.WebId, List.Id, SharePoint.ListItem.Query, SelectedSubFolder == null ? Settings.RootFolderPath : SelectedSubFolder.Path, false);
                var listItems = listItemsXml.Select(li => new SharePoint.ListItem(List, SelectedSubFolder ?? Settings.RootFolder, li));
                foreach (var listItem in listItems)
                {
                    var idString = listItem.Id.ToString(CultureInfo.InvariantCulture);
                    if (cachedListItems.ContainsKey(idString))
                        continue; // happens if an item was added during Update
                    cachedListItems.Add(idString, listItem);
                    yield return new KMOAPICapture.ListItem(idString, listItem.Name);
                }
            }

            internal bool Update(SharePoint.ListItem listItemToSelect = null)
            {
                // get all builtin fields
                var madeChanges = false;
                var appendModeField = (KMOAPICapture.CheckboxField)Form.Fields[AppendModeFieldId];
                var folderField = (KMOAPICapture.ListField)Form.Fields[FolderFieldId];
                var listItemField = (KMOAPICapture.ListFieldEx)Form.Fields[ListItemFieldId];
                var contentTypeField = (KMOAPICapture.ListField)Form.Fields[ContentTypeFieldId];
                var fileNameField = (KMOAPICapture.TextField)Form.Fields[FileNameFieldId];

                // get the append mode
                var prevSelectedAppending = SelectedAppending;
                if (listItemToSelect != null)
                    appendModeField.BoolValue = true;
                SelectedAppending = appendModeField.BoolValue;
                if (firstUpdate || prevSelectedAppending != SelectedAppending)
                {
                    // note the changes
                    madeChanges = true;

                    // set the required and visibility flags of the other fields
                    foreach (var field in fieldMap)
                    {
                        Form.Fields[field.Key].IsRequired = !SelectedAppending && field.Value.IsMandatory;
                        Form.Fields[field.Key].IsHidden = SelectedAppending;
                    }
                    listItemField.IsRequired = SelectedAppending;
                    listItemField.IsHidden = !SelectedAppending;
                    contentTypeField.IsRequired = !SelectedAppending;
                    contentTypeField.IsHidden = SelectedAppending || (List.Flags & SharePoint.ListFlags.EnableContentTypes) == 0;
                }

                // try to get the selected folder
                var prevSelectedSubFolder = SelectedSubFolder;
                var selectedFolderItem = folderField.Items.SingleOrDefault(e => e.Selected);
                if (selectedFolderItem == null)
                {
                    // select the root folder if nothing is selected
                    selectedFolderItem = folderField.Items.Single(e => e.Value == Settings.RootFolderPath);
                    selectedFolderItem.Selected = true;
                }
                SelectedSubFolder = cachedFolders[selectedFolderItem.Value];

                // handle a folder change
                if (firstUpdate || prevSelectedSubFolder != SelectedSubFolder)
                {
                    // note the changes
                    madeChanges = true;

                    // clear the items
                    listItemField.Items.Clear();
                    cachedListItems.Clear();
                    listItemField.RaiseFindEvent = true;

                    // update the content types
                    var folder = SelectedSubFolder ?? Settings.RootFolder;
                    contentTypeField.Items.Clear();
                    foreach (var ct in Settings.FilterContentTypes(folder == null ? List.ContentTypes : folder.ContentTypeOrder))
                        contentTypeField.Items.Add(new KMOAPICapture.ListItem(ct.Id, ct.Name, ct == SelectedContentType));
                }

                // get the selected list item
                var prevSelectedListItem = SelectedListItem;
                var selectedListItem = listItemField.Items.SingleOrDefault(e => e.Selected);
                if (listItemToSelect != null)
                {
                    // check the item
                    if (listItemToSelect.RootFolder != (SelectedSubFolder ?? Settings.RootFolder))
                        throw new ArgumentOutOfRangeException("listItemToSelect");

                    // unselect the current item
                    if (selectedListItem != null)
                        selectedListItem.Selected = false;

                    // check if the given item is already present
                    var idString = listItemToSelect.Id.ToString(CultureInfo.InvariantCulture);
                    if (cachedListItems.ContainsKey(idString))
                    {
                        // find and select it
                        selectedListItem = listItemField.Items.Single(li => li.Value == idString);
                        selectedListItem.Selected = true;
                    }
                    else
                    {
                        // create, select and insert it
                        cachedListItems.Add(idString, listItemToSelect);
                        selectedListItem = new KMOAPICapture.ListItem(idString, listItemToSelect.Name, true);
                        listItemField.Items.Insert(0, selectedListItem);
                    }
                }
                SelectedListItem = selectedListItem == null ? null : cachedListItems[selectedListItem.Value];

                // note the changes
                if (prevSelectedListItem != SelectedListItem)
                    madeChanges = true;

                // get the selected content type
                var prevContentType = SelectedContentType;
                var selectedContentTypeItem = contentTypeField.Items.SingleOrDefault(e => e.Selected);
                if (selectedContentTypeItem == null && contentTypeField.Items.Count > 0)
                {
                    // select the first content type if nothing is selected and there is one
                    selectedContentTypeItem = contentTypeField.Items[0];
                    selectedContentTypeItem.Selected = true;
                }
                SelectedContentType = selectedContentTypeItem == null ? null : List.ContentTypes.Single(ct => ct.Id == selectedContentTypeItem.Value);

                // handle content type changes
                if (firstUpdate || prevContentType != SelectedContentType)
                {
                    // note the changes
                    madeChanges = true;

                    // create a reverse map and remove the old fields from the form
                    var reverseMap = fieldMap.ToDictionary(f => f.Value, f => Form.Fields[f.Key]);
                    foreach (var fieldId in fieldMap.Keys)
                        Form.Fields.Remove(Form.Fields[fieldId]);
                    fieldMap.Clear();

                    // recreate the SharePoint fields
                    if (SelectedContentType != null)
                    {
                        //  and get or create each field
                        foreach (var sharePointField in Settings.FilterFields(SelectedContentType.Fields))
                        {
                            var autoStoreField = (KMOAPICapture.BaseField)null;
                            if (!reverseMap.TryGetValue(sharePointField, out autoStoreField))
                            {
                                autoStoreField = sharePointField.CreateAutoStoreField();
                                var name = sharePointField.InternalName;
                                if (name.Length > 16)
                                {
                                    // find a shorter name
                                    var i = 1;
                                    do
                                    {
                                        var suffix = "." + i++.ToString(CultureInfo.InvariantCulture);
                                        name = sharePointField.InternalName.Substring(0, 16 - suffix.Length) + suffix;
                                    }
                                    while (fieldMap.ContainsKey(name));
                                    List.SnapIn.Component.TaskMessage.Msg(string.Format(Resources.FieldInternalNameShortened, sharePointField.InternalName, name), NSiNetUtil.MsgType.Warning);
                                }
                                autoStoreField.Name = name;
                                autoStoreField.Display = sharePointField.DisplayName;
                                autoStoreField.IsRequired = !SelectedAppending && sharePointField.IsMandatory;
                                autoStoreField.IsHidden = SelectedAppending;
                                autoStoreField.RaiseChangeEvent = true;
                            }
                            fieldMap.Add(autoStoreField.Name, sharePointField);
                            Form.Fields.Add(autoStoreField);
                        }
                    }
                }

                // get the file name
                var prevSelectedFileName = SelectedFileName;
                SelectedFileName = fileNameField.Value.Trim();
                if (firstUpdate || prevSelectedFileName != SelectedFileName)
                    madeChanges = true;

                // reset the first update flag and propagate the updates
                firstUpdate = false;
                if (madeChanges)
                    List.SnapIn.Update(this);
                return madeChanges;
            }
        }

        private readonly Uri webUrl;
        private readonly Guid listId;
        private readonly string webMappedPath;
        private SharePointLists.ListsSoapClient lists;
        private SharePointPeople.PeopleSoapClient people;

        private IEnumerable<XElement> GetContentTypes(SharePoint.List existingList)
        {
            // either fetch a new(er) content type definition or use the existing one
            var listGuid = listId.ToString("B");
            return lists.GetListContentTypes(listGuid, null).Elements(SharePoint.BaseObject.SP + "ContentType").Select(ct => new ContentTypeStub(ct)).Select(stub =>
            {
                if (existingList != null)
                {
                    var existingContentType = existingList.ContentTypes.SingleOrDefault(ect => ect.Id == stub.Id);
                    if (existingContentType != null && existingContentType.Version == stub.Version)
                        return existingContentType.Xml;
                }
                return lists.GetListContentType(listGuid, stub.Id);
            });
        }

        /// <summary>
        /// Gets the capture component.
        /// </summary>
        protected KMOAPICapture.KMCapture Component { get; private set; }

        /// <summary>
        /// Creates a new SharePoint snap-in.
        /// </summary>
        /// <param name="webUrl">The absolute <see cref="Uri"/> of the SharePoint web containing the list.</param>
        /// <param name="webMappedPath">The path to the WebClient-mapped web.</param>
        /// <param name="listId">The <see cref="Guid"/> of the list.</param>
        /// <exception cref="ArgumentNullException"><paramref name="webUrl"/> or <paramref name="webMappedPath"/> is <c>null</c> or <paramref name="listId"/> is <see cref="Guid.Empty"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="webUrl"/> or <paramref name="webMappedPath"/> isn't absolute or not a directory (ie. ending with a forward slash without any query or fragment).</exception>
        protected SnapIn(Uri webUrl, string webMappedPath, Guid listId)
        {
            // check and set the properties
            if (webUrl == null)
                throw new ArgumentNullException("webUrl");
            if (!webUrl.IsAbsoluteUri || webUrl != new Uri(webUrl, "./"))
                throw new ArgumentOutOfRangeException("webUrl");
            if (webMappedPath == null)
                throw new ArgumentNullException("webMappedPath");
            if (!Path.IsPathRooted(webMappedPath))
                throw new ArgumentOutOfRangeException("webMappedPath");
            if (listId == Guid.Empty)
                throw new ArgumentNullException("listId");
            this.webUrl = webUrl;
            this.webMappedPath = Path.GetFullPath(webMappedPath);
            this.listId = listId;
        }

        void KMOAPICapture.ISnapInModule.Initialize(KMOAPICapture.KMCapture component)
        {
            // set the component and the SharePoint services
            Component = component;
            lists = new SharePointLists.ListsSoapClient(CreateBinding(), new EndpointAddress(new Uri(webUrl, "./_vti_bin/Lists.asmx")));
            SetCredentials(lists.ClientCredentials);
            people = new SharePointPeople.PeopleSoapClient(CreateBinding(), new EndpointAddress(new Uri(webUrl, "./_vti_bin/People.asmx")));
            SetCredentials(people.ClientCredentials);
        }

        /// <summary>
        /// Creates a <see cref="Binding"/> to connect with the SharePoint Web Services.
        /// </summary>
        /// <returns>The <see cref="BasicHttpBinding"/> that uses Windows credentials and has a maximum receive limit of 25MB.</returns>
        protected virtual Binding CreateBinding()
        {
            var binding = new BasicHttpBinding(BasicHttpSecurityMode.Transport);
            binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Windows;
            binding.MaxReceivedMessageSize = 25 * 1024 * 1024;
            return binding;
        }

        /// <summary>
        /// Sets the client credentials for the SharePoint Web Services.
        /// </summary>
        /// <param name="credentials">The <see cref="ClientCredentials"/> to adjust.</param>
        protected virtual void SetCredentials(ClientCredentials credentials)
        {
            // allow impersonation
            credentials.Windows.AllowedImpersonationLevel = System.Security.Principal.TokenImpersonationLevel.Impersonation;
        }

        /// <summary>
        /// Intializes a new user session.
        /// </summary>
        /// <param name="form">The <see cref="KMOAPICapture.AsForm"/> used for the session.</param>
        /// <param name="list">The <see cref="SharePoint.List"/> used for the session.</param>
        /// <param name="settings">The session's <see cref="SnapInSettings"/>.</param>
        /// <returns>A user-defined callback value.</returns>
        protected virtual object Initialize(KMOAPICapture.AsForm form, SharePoint.List list, SnapInSettings settings)
        {
            return null;
        }

        /// <summary>
        /// Finalizes a session before submit.
        /// </summary>
        /// <param name="context">The current <see cref="Context"/>.</param>
        /// <param name="fields">All visible <see cref="KMOAPICapture.BaseField"/>s that weren't automatically created.</param>
        /// <param name="values">The final SharePoint field values.</param>
        /// <param name="errorMessage">The error message to be displayed if the result is <c>false</c>.</param>
        /// <returns><c>true</c> if all <paramref name="fields"/> are valid and converted into SharePoint <paramref name="values"/>, <c>false</c> otherwise.</returns>
        protected virtual bool Finalize(Context context, IEnumerable<KMOAPICapture.BaseField> fields, IDictionary<string, string> values, out string errorMessage)
        {
            // make sure that required fields have been set
            var missingField = fields.FirstOrDefault(f => f.IsRequired && string.IsNullOrWhiteSpace(f.Value));
            errorMessage = missingField != null ? string.Format(Resources.RequiredFieldMissing, missingField.Display) : null;
            return missingField == null;
        }

        /// <summary>
        /// Validates a field when its value changed or before submit.
        /// </summary>
        /// <param name="context">The current <see cref="Context"/>.</param>
        /// <param name="field">The field that was changed.</param>
        /// <param name="errorMessage">The description why the field is invalid or <c>null</c>.</param>
        /// <returns><c>true</c> if the field's value is valid, <c>false</c> otherwise.</returns>
        /// <remarks>You must set <see cref="KMOAPICapture.BaseField.RaiseChangeEvent"/> for this method to get called.</remarks>
        protected virtual bool Validate(Context context, KMOAPICapture.BaseField field, out string errorMessage)
        {
            errorMessage = null;
            return true;
        }

        /// <summary>
        /// Retrieves all items of a dynamic list.
        /// </summary>
        /// <param name="context">The current <see cref="Context"/>.</param>
        /// <param name="listField">The <see cref="KMOAPICapture.ListField"/> that is being looked up.</param>
        /// <returns>An enumeration of <see cref="KMOAPICapture.ListItem"/>.</returns>
        protected virtual IEnumerable<KMOAPICapture.ListItem> RetrieveItems(Context context, KMOAPICapture.ListField listField)
        {
            return Enumerable.Empty<KMOAPICapture.ListItem>();
        }

        /// <summary>
        /// Gets suggestions based on the text the user has typed in so far.
        /// </summary>
        /// <param name="context">The current <see cref="Context"/>.</param>
        /// <param name="textField">The <see cref="KMOAPICapture.TextField"/> that invoked auto-completion.</param>
        /// <param name="text">The user input.</param>
        /// <returns>All matching suggestions.</returns>
        protected virtual IEnumerable<string> AutoComplete(Context context, KMOAPICapture.TextField textField, string text)
        {
            return textField.SuggestionList.Where(s => s.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) > -1);
        }

        /// <summary>
        /// Updates custom fields when the <see cref="Context"/> has changed.
        /// </summary>
        /// <param name="context">The current <see cref="Context"/>.</param>
        protected virtual void Update(Context context)
        {
        }

        void KMOAPICapture.ISnapInModule.OnLoad(KMOAPICapture.AsForm form, string userName, KMOAPICapture.UserInfo userInfo)
        {
            // try to find a cached list
            var list = (SharePoint.List)null;
            lock (listCache)
                listCache.TryGetValue(listId, out list);

            // get the current definition and check if it's newer
            var listXml = lists.GetList(listId.ToString("B"));
            if (list == null || new ListStub(listXml).Version != list.Version)
            {
                // parse the definition and add it to the cache
                list = new SharePoint.List(this, listXml, GetContentTypes(list));
                lock (listCache)
                    listCache[listId] = list;
            }

            // get and check the list
            if (list.BaseType != SharePoint.BaseType.DocumentLibrary && (list.BaseType == SharePoint.BaseType.Survey || (list.Flags & SharePoint.ListFlags.DisableAttachments) != 0))
                throw new InvalidOperationException(Resources.ListNotAllowingFiles);

            // get the settings and initalize the custom fields
            var settings = new SnapInSettings(list);
            var state = Initialize(form, list, settings);

            // create the context and build the fields
            Context.Create(form, state, list, settings);
            var context = Context.Get(form);
            context.Update();
        }

        KMOAPICapture.FieldChangeResult KMOAPICapture.ISnapInModule.OnChange(KMOAPICapture.AsForm form, string fieldName, Dictionary<string, string> fieldData)
        {
            // get and update the context
            var context = Context.Get(form);
            context.Update();

            // create the result
            var result = new KMOAPICapture.FieldChangeResult()
            {
                ReturnFormScreenAfterError = false,
                FieldDialogToDisplay = fieldName,
                Result = true,
            };

            // handle only non-builtin fields
            if (!BuiltInFieldIds.Contains(fieldName))
            {
                // find the SharePoint field
                var autoStoreField = form.Fields[fieldName];
                var sharePointField = context.Field(autoStoreField);
                if (sharePointField != null)
                {
                    // try to parse the value and return formatting exceptions
                    try { sharePointField.Parse(autoStoreField); }
                    catch (FormatException e)
                    {
                        result.Result = false;
                        result.ErrorMessage = e.Message;
                    }
                }
                else
                {
                    // use custom validation if enabled
                    if (autoStoreField.RaiseChangeEvent)
                    {
                        var errorMessage = (string)null;
                        result.Result = Validate(context, autoStoreField, out errorMessage);
                        if (!result.Result)
                            result.ErrorMessage = errorMessage;
                    }
                }
            }
            return result;
        }

        List<KMOAPICapture.ListItem> KMOAPICapture.ISnapInModule.OnListSearch(KMOAPICapture.AsForm form, KMOAPICapture.ListField listField)
        {
            // make sure the call is valid
            if (!listField.RaiseFindEvent)
                throw new InvalidOperationException();

            // get the context and retrieve the items
            var context = Context.Get(form);
            var items = Enumerable.Empty<KMOAPICapture.ListItem>();
            switch (listField.Name)
            {
                case FolderFieldId:
                    // add all subfolders
                    items = context.RetrieveFolders();
                    break;

                case ListItemFieldId:
                    // add all items
                    items = context.RetrieveItems();
                    break;

                default:
                    // ignore other builtin fields
                    if (!BuiltInFieldIds.Contains(listField.Name))
                    {
                        // try to find a corresponding SharePoint field and retrieve the items
                        var sharePointField = context.Field(listField);
                        items = sharePointField == null ? RetrieveItems(context, listField) : sharePointField.RetrieveItems().Select(e => new KMOAPICapture.ListItem(e.Key, e.Value));
                    }
                    break;
            }

            // add the items and reset the find property
            var actualListItems = listField is KMOAPICapture.ListFieldEx ? (IList<KMOAPICapture.ListItem>)((KMOAPICapture.ListFieldEx)listField).Items : (IList<KMOAPICapture.ListItem>)listField.Items;
            foreach (var item in items)
                actualListItems.Add(item);
            listField.RaiseFindEvent = false;

            // return the list's items
            // NOTE: the items always have to be present, otherwise KMOAPICapture.KmScanServer.UpdateFieldValues can't set ValueId
            return listField.Items.ToList();
        }

        bool KMOAPICapture.ISnapInModule.OnSubmit(KMOAPICapture.AsForm form, Dictionary<string, string> fieldData, out string errorMessage)
        {
            // get and prepare the context
            var context = Context.Get(form);
            if (context.Update())
            {
                // prevent submit
                errorMessage = Resources.ChangeDuringSubmit;
                return false;
            }

            // check the file name
            if (context.SelectedFileName.Length == 0)
            {
                errorMessage = string.Format(Resources.RequiredFieldMissing, Resources.FileNameField);
                return false;
            }
            if (context.SelectedFileName.IndexOfAny(SnapInSettings.InvalidFileNameChars) > -1)
            {
                errorMessage = string.Format(Resources.InvalidCharsInFileName, string.Join(" ", SnapInSettings.InvalidFileNameChars));
                return false;
            }
            var fileName = context.SelectedFileName;
            if (context.Settings.FileExtension.Length > 0)
                fileName += "." + context.Settings.FileExtension;

            // check the required fields
            if (context.SelectedAppending)
            {
                // check if a list item is selected
                if (context.SelectedListItem == null)
                {
                    errorMessage = string.Format(Resources.RequiredFieldMissing, Resources.ListItemField);
                    return false;
                }
            }
            else
            {
                // check if a content type is selected
                if (context.SelectedContentType == null)
                {
                    errorMessage = Resources.NoMatchingContentTypes;
                    return false;
                }
            }

            // generate and check the root path
            var path = context.List.BaseType == SharePoint.BaseType.DocumentLibrary ?
                context.SelectedSubFolder == null ? context.Settings.RootFolderPath : context.SelectedSubFolder.Path :
                context.List.Path;
            var webPath = context.List.WebPath;
            if (webPath.Length > 1)
                webPath += '/';
            if (!path.StartsWith(webPath, StringComparison.OrdinalIgnoreCase))
                throw new IOException(string.Format(Resources.PathOutsideWeb, path, context.List.WebPath));
            var localPath = Path.Combine(webMappedPath, path.Substring(webPath.Length).Replace('/', Path.DirectorySeparatorChar));

            // add the attachment directory if necessary
            if (context.List.BaseType != SharePoint.BaseType.DocumentLibrary)
            {
                localPath = Path.Combine(localPath, "Attachments");
                path += "/Attachments";

                // skip further processing if we don't know the item yet
                if (!context.SelectedAppending)
                    goto SkipFileExistCheck;

                // append the item path
                var idString = context.SelectedListItem.Id.ToString(CultureInfo.InvariantCulture);
                localPath = Path.Combine(localPath, idString);
                path += "/" + idString;
            }

            // create the complete pathes and check if the file exists
            localPath = Path.Combine(localPath, fileName);
            path += "/" + fileName;
            if (File.Exists(localPath))
            {
                errorMessage = Resources.FileAlreadyExists;
                return false;
            }

        SkipFileExistCheck:

            // collect and check all visible fields
            var sharePointValues = new Dictionary<string, string>();
            var customFields = new List<KMOAPICapture.BaseField>();
            foreach (var autoStoreField in form.Fields.Where(f => !f.IsHidden && !BuiltInFieldIds.Contains(f.Name)))
            {
                // either add the SharePoint value or the custom field
                var sharePointField = context.Field(autoStoreField);
                if (sharePointField != null)
                {
                    // parse the value
                    // NOTE: This might be a wrong value, because if the user presses cancel upon entering an
                    //       invalid value, the wrong value gets applied anyway on the device. The only way to
                    //       remedy this is to set FieldChangeResult.ReturnFormScreenAfterError to true, which
                    //       isn't a nice thing to do, especially if the user enters a long value. 
                    var value = (string)null;
                    try { value = sharePointField.Parse(autoStoreField); }
                    catch (FormatException e)
                    {
                        errorMessage = string.Format(Resources.FieldValueInvalid, autoStoreField.Display, e.Message);
                        return false;
                    }

                    // make sure an empty value is allowed or add the parsed value
                    if (value.Length == 0)
                    {
                        // return an error if a mandatory field is missing
                        if (sharePointField.IsMandatory)
                        {
                            errorMessage = string.Format(Resources.RequiredFieldMissing, sharePointField.DisplayName);
                            return false;
                        }
                    }
                    sharePointValues.Add(sharePointField.InternalName, value);
                }
                else
                {
                    // validate the field and add it to the custom fields (also see the note above)
                    if (autoStoreField.RaiseChangeEvent && !Validate(context, autoStoreField, out errorMessage))
                    {
                        errorMessage = string.Format(Resources.FieldValueInvalid, autoStoreField.Display, errorMessage);
                        return false;
                    }
                    customFields.Add(autoStoreField);
                }
            }

            // finalize the form
            if (!Finalize(context, customFields, sharePointValues, out errorMessage))
                return false;

            // store the list id and append mode
            fieldData[ContextFieldId] = context.List.Id.ToString("B");
            fieldData[AppendModeFieldId] = context.SelectedAppending.ToString(CultureInfo.InvariantCulture);

            // create the item if we're not appending
            if (!context.SelectedAppending)
            {
                // set and store the content type
                sharePointValues.Add("ContentType", context.SelectedContentType.Name);
                fieldData[ContentTypeFieldId] = context.SelectedContentType.Name;

                // handle the different list types
                if (context.List.BaseType == SharePoint.BaseType.DocumentLibrary)
                {
                    // prepare the update operation and store the path
                    sharePointValues.Add("ID", string.Empty);
                    sharePointValues.Add("FileRef", path);
                    fieldData[ListItemFieldId] = path;

                    // touch the file
                    File.WriteAllBytes(localPath, emptyFile);

                    // set the fields
                    try { context.List.SnapIn.UpdateListItem(context.List.Id, context.List.Version, null, SnapInListItemCommand.Update, sharePointValues); }
                    catch
                    {
                        // remove the file upon error
                        try { File.Delete(localPath); }
                        catch { }
                        throw;
                    }
                }
                else
                {
                    // create the new list item
                    var itemXml = context.List.SnapIn.UpdateListItem(context.List.Id, context.List.Version, context.SelectedSubFolder == null ? context.Settings.RootFolderPath : context.SelectedSubFolder.Path, SnapInListItemCommand.New, sharePointValues);
                    var createdItem = new SharePoint.ListItem(context.List, context.SelectedSubFolder ?? context.Settings.RootFolder, itemXml);

                    // store the id and complete the pathes
                    var idString = createdItem.Id.ToString(CultureInfo.InvariantCulture);
                    fieldData[ListItemFieldId] = idString;
                    localPath = Path.Combine(Path.Combine(localPath, idString), fileName);
                    path += "/" + idString + "/" + fileName;

                    // add a pseudo attachment and select the new item
                    lists.AddAttachment(context.List.Id.ToString("B"), idString, fileName, emptyFile);
                    if (context.Settings.AppendMode != SnapInAppendMode.Never)
                        context.Update(createdItem);
                }
            }
            else
            {
                // store an empty content type and the selected item id
                fieldData[ContentTypeFieldId] = string.Empty;
                var idString = context.SelectedListItem.Id.ToString(CultureInfo.InvariantCulture);
                fieldData[ListItemFieldId] = idString;

                // make sure we do the right thing and add a pseudo attachment
                if (context.List.BaseType == SharePoint.BaseType.DocumentLibrary)
                    throw new NotSupportedException();
                lists.AddAttachment(context.List.Id.ToString("B"), idString, fileName, emptyFile);
            }

            // store the directory and file name
            fieldData[FolderFieldId] = Path.GetDirectoryName(localPath);
            fieldData[FileNameFieldId] = Path.GetFileName(localPath);

            // return success
            errorMessage = null;
            return true;
        }

        List<KMOAPICapture.ListItem> KMOAPICapture.ISnapInModule.OnTextSearch(KMOAPICapture.AsForm form, KMOAPICapture.TextField textField, string searchPattern)
        {
            // make sure the call is valid
            if (!textField.IsSuggestionListDynamic)
                throw new InvalidOperationException();

            // get the context
            var context = Context.Get(form);

            // there are no autocomplete builtin fields
            if (BuiltInFieldIds.Contains(textField.Name))
                return new List<KMOAPICapture.ListItem>();

            // either call the SharePoint or the default method
            var sharePointField = context.Field(textField);
            var suggestions = sharePointField == null ? AutoComplete(context, textField, searchPattern ?? string.Empty) : sharePointField.AutoComplete(searchPattern ?? string.Empty);
            return suggestions.Select(s => new KMOAPICapture.ListItem(string.Empty, s)).ToList();
        }

        List<KMOAPICapture.AsTreeNode> KMOAPICapture.ISnapInModule.OnTreeSearch(KMOAPICapture.AsForm form, KMOAPICapture.TreeField treeField, KMOAPICapture.AsTreeNode currentNode)
        {
            // not used, perform the default operation
            return currentNode == null ?
                treeField.Nodes.ToList() :
                currentNode.Nodes.ToList();
        }

        /// <summary>
        /// Searches for SharePoint principals that match a given user name.
        /// </summary>
        /// <param name="userName">The logon or display user name.</param>
        /// <param name="includeGroups">Indicates whether the <paramref name="userName"/> may also refer to a group.</param>
        /// <returns>The top 7 <see cref="SharePointPeople.PrincipalInfo"/> matches.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="userName"/> is <c>null</c>.</exception>
        public IEnumerable<SharePointPeople.PrincipalInfo> SearchPrincipals(string userName, bool includeGroups)
        {
            // check the input, calculate the limit and return the results from SharePoint
            if (userName == null)
                throw new ArgumentNullException("userName");
            var limitTo = SharePointPeople.SPPrincipalType.User;
            if (includeGroups)
                limitTo |= SharePointPeople.SPPrincipalType.SecurityGroup | SharePointPeople.SPPrincipalType.SharePointGroup;
            Component.TaskMessage.Msg("Calling SearchPrincipals:", NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    Text: " + userName, NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    Type: " + limitTo.ToString(), NSiNetUtil.MsgType.Trace);
            return people.SearchPrincipals(userName, 7, limitTo);
        }

        /// <summary>
        /// Resolves user logon names into SharePoint principals.
        /// </summary>
        /// <param name="userNames">The names of users to resolve.</param>
        /// <param name="includeGroups">Indicates whether <paramref name="userNames"/> may also include group names.</param>
        /// <returns>An array of <see cref="SharePointPeople.PrincipalInfo"/> corresponding to the <paramref name="userNames"/> entry with the same index.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="userNames"/> is <c>null</c>.</exception>
        public SharePointPeople.PrincipalInfo[] ResolvePrincipals(string[] userNames, bool includeGroups)
        {
            // check the input
            if (userNames == null)
                throw new ArgumentNullException("userNames");
            if (userNames.Length == 0)
                return new SharePointPeople.PrincipalInfo[0];

            // calculate the limit
            var limitTo = SharePointPeople.SPPrincipalType.User;
            if (includeGroups)
                limitTo |= SharePointPeople.SPPrincipalType.SecurityGroup | SharePointPeople.SPPrincipalType.SharePointGroup;

            // turn the enum into an array
            var names = new SharePointPeople.ArrayOfString();
            names.AddRange(userNames);

            // resolve the principals
            Component.TaskMessage.Msg("Calling ResolvePrincipals:", NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    Keys: " + string.Join("; ", userNames), NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    Type: " + limitTo.ToString(), NSiNetUtil.MsgType.Trace);
            return people.ResolvePrincipals(names, limitTo, true);
        }

        /// <summary>
        /// Creates, updates or deletes a SharePoint item.
        /// </summary>
        /// <param name="listId">The <see cref="Guid"/> of the list containing the item.</param>
        /// <param name="listVersion">The list schema version or <c>0</c>.</param>
        /// <param name="folder">The root folder where to create items or <c>null</c>.</param>
        /// <param name="command">The command to perform.</param>
        /// <param name="values">The field values.</param>
        /// <returns>The row of the <see cref="SnapInListItemCommand.New"/> item.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="listId"/> is <see cref="Guid.Empty"/> or <paramref name="values"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="listVersion"/> is less than 0 or <paramref name="values"/> is empty.</exception>
        /// <exception cref="Win32Exception">An exception occured when performing the operation on the server.</exception>
        public XElement UpdateListItem(Guid listId, int listVersion, string folder, SnapInListItemCommand command, IDictionary<string, string> values)
        {
            // check the input arguments
            if (listId == Guid.Empty)
                throw new ArgumentNullException("listId");
            if (values == null)
                throw new ArgumentNullException("values");
            if (values.Count == 0)
                throw new ArgumentOutOfRangeException("values");
            if (listVersion < 0)
                throw new ArgumentOutOfRangeException("listVersion");

            // create the batch
            var method =
                new XElement("Method",
                    new XAttribute("ID", 1),
                    new XAttribute("Cmd", command));
            method.Add(values.Select(v => new XElement("Field", new XAttribute("Name", v.Key), v.Value)).ToArray());
            var batch =
                new XElement("Batch",
                    new XAttribute("OnError", "Continue"),
                    new XAttribute("ListVersion", listVersion),
                    new XAttribute("Properties", true),
                    method);
            if (folder != null)
                batch.Add(new XAttribute("RootFolder", folder));

            // perform the operation and return the row
            Component.TaskMessage.Msg("Calling UpdateListItems:", NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    ListId: " + listId.ToString("B"), NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    Batch: " + batch.ToString(SaveOptions.DisableFormatting), NSiNetUtil.MsgType.Trace);
            var result = lists.UpdateListItems(listId.ToString("B"), batch).Element(SharePoint.BaseObject.SP + "Result");
            var errorCodeString = result.Element(SharePoint.BaseObject.SP + "ErrorCode").Value;
            var errorCode = errorCodeString.StartsWith("0x", StringComparison.Ordinal) ? int.Parse(errorCodeString.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) : int.Parse(errorCodeString, NumberStyles.Integer, CultureInfo.InvariantCulture);
            if (errorCode != 0)
            {
                // throw the error
                var errorMessage = result.Element(SharePoint.BaseObject.SP + "ErrorText");
                if (errorMessage == null || string.IsNullOrWhiteSpace(errorMessage.Value))
                    throw new Win32Exception(errorCode);
                else
                    throw new Win32Exception(errorCode, errorMessage.Value);
            }
            return result.Element("{#RowsetSchema}row");
        }

        /// <summary>
        /// Queries a list for items.
        /// </summary>
        /// <param name="webId">The <see cref="Guid"/> of the SharePoint web that contains the list.</param>
        /// <param name="listId">The <see cref="Guid"/> of the list to query.</param>
        /// <param name="query">The query in CAML.</param>
        /// <param name="fields">The name of fields to return or <c>null</c> to query mandatory and metadata fields.</param>
        /// <returns>An enumeration of all return rows.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="webId"/> or <paramref name="listId"/> is <see cref="Guid.Empty"/>.</exception>
        public IEnumerable<XElement> GetListItems(Guid webId, Guid listId, XElement query, params string[] fields)
        {
            return GetListItems(webId, listId, query, null, true, fields);
        }

        /// <summary>
        /// Queries a list for items.
        /// </summary>
        /// <param name="webId">The <see cref="Guid"/> of the SharePoint web that contains the list.</param>
        /// <param name="listId">The <see cref="Guid"/> of the list to query.</param>
        /// <param name="query">The query in CAML.</param>
        /// <param name="folder">The path to the folder where to start the query or <c>null</c> to start at the root of the list.</param>
        /// <param name="recursive">Indicates whether subfolders should be queried as well.</param>
        /// <param name="fields">The name of fields to return or <c>null</c> to query mandatory and metadata fields.</param>
        /// <returns>An enumeration of all return rows.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="webId"/> or <paramref name="listId"/> is <see cref="Guid.Empty"/>.</exception>
        public IEnumerable<XElement> GetListItems(Guid webId, Guid listId, XElement query, string folder, bool recursive, params string[] fields)
        {
            // check the input arguments
            if (webId == Guid.Empty)
                throw new ArgumentNullException("webId");
            if (listId == Guid.Empty)
                throw new ArgumentNullException("listId");

            // build the view fields (if no fields are given then expand the meta data)
            var metaView = fields == null || fields.Length == 0;
            var viewFields = new XElement("ViewFields");
            if (metaView)
                viewFields.Add(new XAttribute("Properties", true));
            else
                viewFields.Add(fields.Select(f => new XElement("FieldRef", new XAttribute("Name", f))).ToArray());

            // build the query options
            var queryOptions =
                new XElement("QueryOptions",
                    new XElement("IncludeMandatoryColumns", metaView.ToString(CultureInfo.InvariantCulture)));
            if (folder != null)
                queryOptions.Add(new XElement("Folder", folder));
            if (recursive)
                queryOptions.Add(new XElement("ViewAttributes", new XAttribute("Scope", "RecursiveAll")));

            // execute the query
            Component.TaskMessage.Msg("Calling GetListItems:", NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    ListId: " + listId.ToString("B"), NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    WebId: " + webId.ToString("B"), NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    Query: " + query.ToString(SaveOptions.DisableFormatting), NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    ViewFields: " + viewFields.ToString(SaveOptions.DisableFormatting), NSiNetUtil.MsgType.Trace);
            Component.TaskMessage.Msg("    QueryOptions: " + queryOptions.ToString(SaveOptions.DisableFormatting), NSiNetUtil.MsgType.Trace);
            return lists.GetListItems
            (
                listId.ToString("B"),
                null,
                query,
                viewFields,
                int.MaxValue.ToString(CultureInfo.InvariantCulture),
                queryOptions,
                webId.ToString("B")
            ).Element("{urn:schemas-microsoft-com:rowset}data").Elements("{#RowsetSchema}row");
        }
    }
}
