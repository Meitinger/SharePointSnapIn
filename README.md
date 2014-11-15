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
- Content-type awareness: The user can specify what type of document or list
  item they want to create, and the input fields are adjusted accordingly.
- The settings can be fine-tuned based on the logged-on user, see the settings
  section for more detail.
- (for *developers* and users) It contains an advanced list field that allows
  unlimited, dynamic multiple selections, localized empty selection strings and
  an option to clear the selection.

To upload the actual file to SharePoint it uses the *Send to Folder* filter or
route component in combination with the Windows WebDAV Redirector service. The
following section describes how to set up the library and the redirector.


Setup
-----
First build the library. Depending on your environment, you need to change the
library references to point to the AutoStore libraries (they are located in the
`%ProgramFiles(x86)%\NSi\AutoStore Foundation 6` folder).

Now make sure the WebDAV redirectory is installed and configured properly. If
you're running Windows Server, install the Desktop Experience feature.
If you're accessing SharePoint via a FQDN, make sure the client credentials are
forwarded properly, see http://support.microsoft.com/kb/943280.

Next copy the library to `%ProgramFiles(x86)%\NSi\AutoStore Foundation 6`,
launch AutoStore Process Designer, create a *Basic Form* (consult the AutoStore
documentation for help) and click on the *Create / Edit Snap-In* button.
Click on *References* and add the previously copied library path to the list.
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

The list referenced by the GUID must be located directly under the given site.
Also make sure that *Form is loaded* and *Form is submitted* are checked.

The final step is to configure the *Send to Folder* filter or route component.
To do this, add a path entry with the following settings:
- Folder Path: `~KMO::%SP.FOLDER%~`
- Overwrite Existing File: `checked`
- Rename File: `checked`
- Schema: `~KMO::%SP.FILENAME%~`
- Replace Invalid characters with "_": `unchecked`


Settings and Advanced Usage
---------------------------
(coming soon, in the meantime look at the documentation in the source)
