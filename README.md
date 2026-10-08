# PDF TOC Extractor
This simple script attempts to capture the table of contents from all *PDF* files in a folder.

## The Problem

Occasionally, you just need a list of chapters in a set of books, report or anthologies.

## Features

* Captures the table of contents from the file *bookmarks*.
* If the file doesn't have bookmarks, then the script tries to find and read the table of contents text.
* Bulk operation on all *PDF* and *EPUB* files in a folder.

## Usage

### Preparation

Before running the script, store all the files you wish the script to read into a single folder.

### Running the script

When you run the *Python* script, it will pop up a file explorer/finder window asking you to select the folder housing the source *PDF* files.

The script will process the *PDF* and *EPUB* files one by one. 

If the file contains *Bookmarks*, then the script will capture the list of chapters based on that. If not, then the script will try to find and read the table of contents text.

### Output

The script will save the output as an *Excel* file in the same folder as the original.

### Note

1. The script uses some standard *Python* libraries. If you don't have them installed on your system, then in the first run, the script will try to install these dependencies. If the script can't install these dependencies, for instance if your PC environment precludes it, then it will usually give you the console commands you can use to install these.
2. If some of your libraries and executables sit outside the *Path*, such as if you don't have Admin rights to your work laptop, then you should include the folder addresses in the '*path.txt*' file, which should sit in the same folder as the '*pdf-toc-extractor.py*' file, e.g. '*C:\Users\Username\AppData\Roaming\Python\Python313\Scripts*' and '*C:\Users\Username\AppData\Roaming\Python\Python313\site-packages*'.

## Caveat

Tested on *Windows 11 Education 64-bit*.

Not tested on *Apple iOS* or *Linux*.

## Never run a Python script before?

It's straightforward, but you may need to install *Python* on your machine first.

### Install Python

*Anaconda* is one of the most popular distributions of *Python*. Download and install from https://www.anaconda.com/download

Installation is simple, but if you need help, check out https://www.anaconda.com/docs/getting-started/anaconda/install/overview

### Start Spyder

*Anaconda* comes with *Spyder IDE*. Start *Spyder*.

Once *Spyder* is ready, open the file '*pdf-toc-extractor.py*' that has the script.

All that's left is for you to hit 'Run', i.e. the green 'Play' button.

## Apps without downloading Anaconda / Python

*Windows* and *MacOS* executables (apps) are available to download under *Releases* - always choose the newest version.

Because I am not a software company with an expensive *Code Signing Certificate* when you download the executable, and again when you try to run it, your computer will warn you not to, because it comes from an unverified source. Based on your OS, it may not always be obvious how to bypass the warnings. 