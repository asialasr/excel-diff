# excel-diff
## Overview
This project is a (currently rather shoddy) diff-tool for excel files.  It isn't particularly well tested at the moment, but it is able to detect added & deleted sheets, modified sheets, and modified cells.  It takes as input the path to two excel files and outputs a single excel file which is formatted to show added, deleted, and modified cells and uses prefixes on the tab names to indicate new, deleted, and modified tabs.

I haven't gotten a chance to work on this in a while, but hope to get back to it soon to clean it up and test it better.

## Libraries
This project uses the Python libraries xlsxwriter and xlrd (can be installed via pip), but I think that's it.  I'll go through and double-check that all libraries are listed.

// TODO(asialasr): double-check libraries

## Python
I'm going off of memory, but I think that I was using Python36 for this, but I'll also try to double-check that.

// TODO(asialasr): add Python minimum version

## User Interface
What started me getting off of working on this was that I started to work on a user interface for this, which in a roundabout way led to me getting distracted in trying to learn OpenGL and caused this to fall by the wayside.  I do also have a really awful UWP interface for this, but it's so bad that it would be detrimental to put it out there, even though it mostly works.

I'll update this whenever I get back around to creating the user interface, which I hope to do soon, then I'll get started on cleaning all of this up too.

// TODO(asialasr): update when user interface is implemented
