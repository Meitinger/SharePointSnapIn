SharePoint Snap-In for AutoStore
================================


Description
-----------
This library is used together with NSi AutoStore and the Konica Minolta Capture
component. It provides an alternative implementation to the SharePoint form and
filter/route component with the following advantages:
- Input fields are just as they are on SharePoint, meaning that lookup columns
  show the choices, numeric columns are properly parsed and checked, user and
  group fields allow searching for SharePoint principals, and so on.
- When lists are used instead of document libraries, this library allows users
  not only to create new list items but also append attachments to existing
  ones.
- Content-type awareness: Users can specify what type of document or list item
  they want to create, and the input fields are adjusted accordingly.
- The settings can be fine-tuned based on the logged-on user, see the last
  section for more details.
- It uses SharePoint's web services directly, thus eliminating the need for a
  separate feature or web service to be installed on the SharePoint server and
  also (hopefully) ensuring compatibility with future SharePoint versions.
- (for **developers** and users) It contains an advanced list field that allows
  unlimited, dynamic multiple selections, localized empty selection strings and
  an option to clear the selection.

To upload the actual file to SharePoint it uses the *Send to Folder* filter or
route component in combination with the Windows WebDAV Redirector service. The
following section describes how to set up the library and the redirector.


Build and Install
-----------------
First build the library. Depending on your environment, you need to change the
library references to point to the AutoStore libraries (they are located in the
`%ProgramFiles(x86)%\NSi\AutoStore Foundation 6` folder).

Now make sure the WebDAV redirector is installed and configured properly. If
you're running Windows Server, install the *Desktop Experience* feature.
If you're accessing SharePoint via a FQDN, make sure the client credentials are
forwarded properly, see http://support.microsoft.com/kb/943280.
You can test it by running the following command as your AutoStore batch user:
    echo "Hello World!" > \\<sharepoint-server-fqdn>[@SSL[@<port>]]\DavWWWRoot\<site>\<doklib>\test.txt

Next copy the library to `%ProgramFiles(x86)%\NSi\AutoStore Foundation 6`,
launch AutoStore Process Designer, create a 'Basic Form' (consult the AutoStore
documentation for help) and click on the 'Create / Edit Snap-In' button.
Click on 'References' and add the previously copied library path to the list.
Now copy the following snippet into the script, replace the two strings and the
GUID with the appropriate values, save the script and close the editor.

    using System;
    using SharePointSnapIn;
    
    namespace SharePointSample
    {
        public class SampleSnapIn : SnapIn
        {
            public SampleSnapIn() : base(
                new Uri("https://sharepoint.example.com/<your>/<web>/<site>/"),
                @"\\sharepoint.example.com@SSL\DavWWWRoot\<your>\<web>\<site>\",
                new Guid("{YOUR-LIST-GUID-HERE}")) {}
        }
    }

The list referenced by the GUID **must** be located in site the path points to.
Also, make sure that 'Form is loaded' and 'Form is submitted' are checked.

The final step is to configure the *Send to Folder* filter or route component.
To do this, add a path entry with the following settings:
- Folder Path: `~KMO::%SP.FOLDER%~`
- Overwrite Existing File: `checked`
- Rename File: `checked`
- Schema: `~KMO::%SP.FILENAME%~`
- Replace Invalid characters with "_": `unchecked`


Settings and Advanced Usage
---------------------------
There are a couple of `protected virtual` methods in the `SnapIn` class that
can be overridden to customize its behavior and extend its functionality.
All of them are documented in the source code, but usually the only method you
will want to change is `Initialize` unless you want to implement custom fields,
which is covered in the second example.
But first let's look at the following example of a monthly expense report form:

    protected override object Initialize(KMOAPICapture.AsForm form, SharePointSnapIn.SharePoint.List list, SharePointSnapIn.SnapInSettings settings)
    {
        settings.RootFolderPath = "Expenses " + DateTime.Today.Year;
        settings.AllowSubFolders = false;
        settings.DefaultFileName = form.User.UserName + " - " + DateTime.Today.Month;
        return null;
    }

This places the scanned report into the folder of the current year (without the
user being able to change it) and pre-fills the file name with the currently
logged on user name and the month. There are of course many more setting that
can be adjusted, just have a look at the documentation of `SnapInSettings`.

But what if you want to handle a field yourself? For our next and final example
let's pretend you have a task list and you want to make sure that the field
*Assigned To* can only be assigned a person that directly reports to the logged
on user:

    protected override object Initialize(KMOAPICapture.AsForm form, SharePointSnapIn.SharePoint.List list, SharePointSnapIn.SnapInSettings settings)
    {
        // tell the snap-in to ignore the field
        settings.IgnoreField("AssignedTo");

        // add our own dynamic list field
        form.Fields.Add(new KMOAPICapture.ListFieldEx()
        {
            Name = "MyAssignedTo",
            Display = list.Fields.Single(f => f.InternalName == "AssignedTo").DisplayName,
            IsRequired = true,
            RaiseFindEvent = true,
            AllowMultipleSelection = false,
        });
        return null;
    }

    protected override void Update(Context context)
    {
        // only show the field if the user is not appending an existing item
        context.Form.Fields.Single(f => f.Name == "MyAssignedTo").IsHidden = context.SelectedAppending;
    }

    private string[] GetLogonNamesOfDirectReportsFromUser(KMOAPICapture.UserInfo user)
    {
        // the implementation if this method is beyond the scope of this example
        ...
    }

    protected override IEnumerable<KMOAPICapture.ListItem> RetrieveItems(Context context, KMOAPICapture.ListField listField)
    {
        // if it's not our field forward the call to the base method
        if (listField.Name != "MyAssignedTo")
            return base.RetrieveItems(context, listField);
            
        // resolve the logon names and return their ids and display names
        var principals = ResolvePrincipals(GetLogonNamesOfDirectReportsFromUser(context.Form.User), false);
        return principals.Where(p => p.IsResolved).Select(p => new KMOAPICapture.ListItem(p.UserInfoID + ";#", p.DisplayName));
    }

    protected override bool Finalize(Context context, IEnumerable<KMOAPICapture.BaseField> fields, IDictionary<string, string> values, out string errorMessage)
    {
        // perform the mandatory field check
        if (!base.Finalize(context, fields, values, out errorMessage))
            return false;

        // add the value if the user is not appending
        if (!context.SelectedAppending)
            values["AssignedTo"] = fields.Single(f => f.Name == "MyAssignedTo").Value;
        errorMessage = null;
        return true;
    }

The above code creates a list field that is filled with all direct subordinates
and that replaces the SharePoint field.
If instead you wanted the user to just have his or her subordinates as
suggestion, you could do something like this:

    private class State
    {
        public SharePointSnapIn.SharePointPeople.PrincipalInfo ResolvedPrincipal { get; set; }
    }

    protected override object Initialize(KMOAPICapture.AsForm form, SharePointSnapIn.SharePoint.List list, SharePointSnapIn.SnapInSettings settings)
    {
        // tell the snap-in to ignore the field
        settings.IgnoreField("AssignedTo");

        // add our own autocomplete text field with validation
        form.Fields.Add(new KMOAPICapture.TextField()
        {
            Name = "MyAssignedTo",
            Display = list.Fields.Single(f => f.InternalName == "AssignedTo").DisplayName,
            IsRequired = true,
            IsSuggestionListDynamic = true,
            SuggestionListType = KMOAPICapture.TextSuggestionListType.List,
            Value = string.Empty,
			RaiseChangeEvent = true,
        });
        return new State();
    }

    protected override void Update(Context context)
    {
        // only show the field if the user is not appending an existing item
        context.Form.Fields.Single(f => f.Name == "MyAssignedTo").IsHidden = context.SelectedAppending;
    }

    private string[] GetLogonNamesOfDirectReportsFromUser(KMOAPICapture.UserInfo user)
    {
        // the implementation if this method is beyond the scope of this example
        ...
    }

    protected override IEnumerable<string> AutoComplete(Context context, KMOAPICapture.TextField textField, string text)
    {
        // return the logon names if it's our field or forward the call
        return textField.Name == "MyAssignedTo" ? GetLogonNamesOfDirectReportsFromUser(context.Form.User) : base.AutoComplete(context, textField, text);
    }

    protected override bool Validate(Context context, KMOAPICapture.BaseField field, out string errorMessage)
    {
        // forward the call to the base method if it isn't our field
        if (field.Name != "MyAssignedTo")
            return base.Validate(context, field, out errorMessage);

        // get the state and clear the old principal
        var state = (State)context.UserState;
        state.ResolvedPrincipal = null;

        // if something was entered try to resolve the value into a principal
        if (field.Value.Length > 0)
        {
            state.ResolvedPrincipal = ResolvePrincipals(new string[] { field.Value }, false)[0];
            if (!state.ResolvedPrincipal.IsResolved)
            {
                // notify the user and abort the operation if no matching principal was found
                errorMessage = "The given principals cannot be resolved.";
                return false;
            }
            
            // replace the value with the account name in case only a partial name was entered
            field.Value = state.ResolvedPrincipal.AccountName;
        }

        // return success
        errorMessage = null;
        return true;
    }

    protected override bool Finalize(Context context, IEnumerable<KMOAPICapture.BaseField> fields, IDictionary<string, string> values, out string errorMessage)
    {
        // perform the mandatory field check
        if (!base.Finalize(context, fields, values, out errorMessage))
            return false;

        // add the value if the user is not appending
        if (!context.SelectedAppending)
            values["AssignedTo"] = ((State)context.UserState).ResolvedPrincipal.UserInfoID + ";#";
        errorMessage = null;
        return true;
    }

Here a user state object in combination with `Validate` is used to keep the
actual principal separate from the field's value. Note that the check for an
empty value is not done in `Validate` but by calling the base method in
`Finalize`, the latter being only called if `Validate` succeeds.

The last thing to mention is that if you need to update or get SharePoint list
items in your code, you can make use of `UpdateListItem` and `GetListItems`.
