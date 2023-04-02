## Introduction
For GIT (Github) testing only! Tool for injecting path into a existing docx files. You can use this too for capturing NTLM hashes.
## Current Releases
0.1 - Initial commit. <br />
## Platforms
Version `v0.1` is semi-automatic version for __Windows__ hosts only. __Microsoft Office Word__ required!
## Usage
Just Launch it with parameters:
 > injector.py -d `directory` -s `server` -u `url>` [-r] <br />

Parameters:<br />
__-d (--directory)__: Directory to scan files. Can be folder on your computer or network drive; For example: _C:\\Users\Admin\Documents_, relative path or single document;<br />
__-s (--server)__: Server where you can launch your SMB poisoner. For example: __192.168.122.2\Template.docx__ or __\\192.168.122.2\Template.docx__. You can use ['Responsibility Tool'](https://github.com/dardepin/responsibility) as poisoner;<br />
__-u (--url)__: Any url with image to insert. Insert this url to insert image inside documents;<br />
__-r (--recursive)__: Recursive option. Does not affect when directory is a single document;

Temporary document(s) will be opened in __Word__ application. You must to change it (them) manually. To link an image, open the insert tab and click the Pictures icon. This will bring up the explorer window. In the file name field, enter the URL and hit the insert drop down to choose “Link to File”. I recommend insert the image in a place that is less likely to be changed or deleted. Once linked, the image can be sized down to nothing. Make sure to save the changes to the document.

Alternatively you can do all this actions manually, like in this article: ['Microsoft Word – UNC Path Injection with Image Linking'](https://www.netspi.com/blog/technical/network-penetration-testing/microsoft-word-unc-path-injection-image-linking/), but this tool's advantage is that it tries to keep metadata (like a file creation/modification times, document authors and some other) and your time saving.
## Development plans
I will try to add image inserting automation in the next version
## Licenses
Use and modify on your own risk.
