# Import Photos
An attempt to create an app that copies photos from an iPhone into a custom directory structure using python

Currently doesn't actually copy files, beacuse pywin32 has difficulty working with the special folders that the iPhone presents to windows. I need to try again with newer versions of pywin32.

Requires pywin32 (pre compiled binary avaliable at http://www.lfd.uci.edu/~gohlke/pythonlibs/#pywin32)<br/>
and optionally exifread, for date validation (https://pypi.python.org/pypi/ExifRead)
