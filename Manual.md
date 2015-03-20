# Introduction #

On the Dune Streamer a Media Index can be created by a system of folders and files.
Even for a small number of files, it is a lot of work to create and maintain such
an index. Various applications have been created for Video (yaDIS, Zappiti, etc.).
For Music only Dune Music Library (promising, but not finished) and Musicnizer
(not free) exist.

This script uses an existing MediaMonkey library and creates a simple Music Index.
The Menu can be searched by: Album, Artist And Year. Below these are subfolders
to make the catalog manageable and finding music easy and quick.

The Index is created with the Android App DMC in mind. DMC can read the Dune Index
folder and present the same menu on your phone/tablet. This means, it is possible
to navigate through your music without having the TV on and without pointing your
remote to the TV (as it works with wifi).

This is my first script in MediaMonkey. I wanted a simple Menu Index, which it is,
and it works very well for me. If you are looking for a more extensive index,
try Dune Music Library or Musicnizer. I'd thought to share it at this moment. Use it,
like it, change it, it's yours.


# Installation #

Not a fancy OneClickInstallation(c), but a few steps are needed. Since version 1.3, the Main Index Folder is separated from the actual MediaMonkey script. You can still download a predefined Index Folder, or create one yourself by using the batch files included (ImageMagick).

## Script ##

The script itself must be copied to the MediaMonkey Folder:
  1. Unzip the file to a temporary folder
  1. Copy the content of the folder scripts to your local MediaMonkey Scripts folder
  1. Change your Scripts.ini file. An example is enclosed below.

## Index Folder ##

To create the basic Index Folder, you have two options (A: use a default or B: create by using batch files).

### A: Use the default (download 10Mb) ###
  1. download the file
  1. unpack
  1. copy the DuneIndex Folder to your Dune Media Player

### B: Create these by running a DOS batch script (needs ImageMagick) ###
  1. goto subfolder CreateIndexFolder
  1. edit `_HitMe_CreateDuneIndex.cmd` and fill in the ImageMagick folder
  1. start `_HitMe_CreateDuneIndex.cmd` and wait for around 80s (on my Core2Duo 2.66Ghz)
  1. copy the DuneIndex Folder to your Dune Media Player

This DuneIndex folder is the Music Index Folder. Name it whatever you like. Change the first lines of the script DuneCatalog.vbs to your network/folder names and locations.

## Settings in DuneCatalog.vbs ##

In the GUI the settings can be made to find the music files and the location to write the Index to. To make these settings permanent, it is possible to change these in the script. It helps a lot to have some knowledge of the ["dune\_folder.txt mechanism"](http://dune-hd.com/support/misc/dune_folder_howto.txt) to understand these settings. Here is a screenshot to explain some folders and locations.

![http://dunecatalog.googlecode.com/svn/wiki/images/GUI-1.png](http://dunecatalog.googlecode.com/svn/wiki/images/GUI-1.png)

The important settings are:
| **name in script** | **value in example** | **explanation** |
|:-------------------|:---------------------|:----------------|
| `DuneDriveLetter` | `J` | Drive Letter of the Dune in Windows |
| `DuneMusicFolderName` | `storage_name://DuneHDD/` | Location of the local (Dune) Music files. |
| `NetworkDriveLetter` | `U` | Drive Letter of the network music path in Windows |
| `NetworkMusicFolderName` | `smb://bat/music/` | Location of the network Music files, as seen by the Dune. |
| `DuneIndexFolder` | `J:\_index\music\` | Location of the music index on the Dune Player |

The two drive letters in rows 1 and 3 of the table are the easiest to understand. The music files on these drives can be indexed. Anything on another drive will not be indexed.

In the example, the Dune harddisk share (`\\DUNE\DuneHDD`) is mapped to `J:`. One of the possibilities for the Dune to see music on its local drive is using `//storage_name/<share>`, so row 2 in this example is: `storage_name://DuneHDD/`.

The network folder share (`\\bat\music\`) is mapped to `U:` This network location is accessed by the smb protocol. The Dune can see this with the next setting: `smb://bat/music/`, see row 4.

The location of the index (row 5) is on the Dune at `J:\_index\music\`. It is also possible to create this on a local drive as well, which will speed up the creation of the index. The index can be moved to the Dune later.
<br><b>N.B. In v1.4 and earlier, this Index Location must end with a backslash. This is not needed in later versions.</b>

<h1>Usage</h1>

<ol><li>Make sure your MediaMonkey Library is up-to-date<br>
</li><li>Select one or more albums in MediaMonkey<br>
</li><li>Start the Script & press OK. The Index files will be generated and copied to the Dune Index Folder<br>
</li><li>Open the Index Folder on your Dune, select your music and enjoy</li></ol>


<h1>Features</h1>

<ul><li><b>Sorting</b> Selected Files will be sorted by Album, then by Disc.Track Number (on/off with checkbox).<br>
</li><li><b>Collections</b> Albums that have multiple artists and the word "various" as the first word in the 'albumartist' tag are grouped together as "various" albums.<br>
</li><li><b>Simple Layout</b> The Menu Layout on Dune is simple: easy access and very accessible by DMC<br>
</li><li><b>Local and Network Music Source</b> One local and one network music source can be indexed by this script and accessed by the Dune</li></ul>


<h1>Warning</h1>

Every new Index File will overwrite existing Index Files (is it a bug or a feature?), so be aware.<br>
<br>
<br>
<h1>Scripts.ini (example)</h1>

<pre><code>[DuneCatalog]<br>
Filename=DuneCatalog.vbs<br>
Procname=OnStartUp<br>
Order=99<br>
DisplayName=Dune Catalog Creator<br>
Description=Creates a Music Index for the Dune Media Player<br>
Language=VBScript<br>
ScriptType=0<br>
</code></pre>


<h1>Needed</h1>

<ul><li>HDI Dune Media Streamer<br>
</li><li>MS Windows<br>
</li><li>MediaMonkey<br>
</li><li>Network access to the Dune, with a network map name<br>
</li><li>up-to-date MediaMonkey file library</li></ul>


<h1>Tested on/with</h1>

<ul><li>Dune Smart D1<br>
</li><li>Windows 7<br>
</li><li>MediaMonkey v4.06.1501</li></ul>


<h1>Limitations</h1>

DuneCatalog does not do:<br>
<ul><li>Album Art Scraping<br>
</li><li>Mass Tagging<br>
</li><li>Cover art can not be png (limitation by VBScript). jpg, bmp, gif is allowed