# VBSUBCLASS
Robust Subclassing for use in the VB6 IDE

* All Visual Basic
* No Assembly Dispatching Stub, or Self Modifying code
* Uses Windows subclassing API (comctl32.dll), no need to track what's been subclassed
* Checks IDE Break Mode before dispatching (vba6.dll#EbMode).

Breaking in the Subclassed WindowProc rarely crashes.<br>
Stopping and Resetting the Project will still crash the IDE. (don't do it)
Don't use End either.
