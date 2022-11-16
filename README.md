# Outlook Archiver

This is a tool created to archive/save locally selected outlook emails as msg files along with each email's attachments.
This tool was initially created for a client but can be adapted for general usage.

Open the tool and outlook > Highlight the emails you would like to archive > Press Export

# Default Export Path

1. The tool's export folder path defaults to the current working directory i.e. the folder the app was launched in
2. If the "archived-mail" folder does not exist within the default folder it will prompt the user asking if a folder should be created
3. Any archived emails will then be sent to this "archived-mail" folder or if the user wishes to not create one it will instead export to just the default path
4. The user is also able to define their own custom path in which case, the tool will not prompt for the creation the "archived-mail" folder

# Creation of folders by the tool

1. This tool creates a few folders during its operation outside of the "archived-mail" folder described in the previous section.
2. This tool creates folders for each category defined in the tool. For example, if there are 3 categories defined, 3 folders will be created for each category in the selected export folder.
3. Any archived email will be saved in the subfolder corresponding to its defined category.

# Logging

1. This tool creates a folder , "logs", in the current working directory.
2. In this folder, the tool creates a log file called "app.log"
3. Refer to this file if there are errors in the tool. This file also provides an audit trail for what and when something was archived.
