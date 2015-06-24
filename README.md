# DXSUBCLASS
Robust Subclassing DLL for debugging in the VB6 IDE

* All Visual Basic
* No Assembly Dispatching Stub, or Self Modifying code, at the cost of a single subclass per hWnd
* Uses Windows subclassing API (comctl32.dll), no need to track what's been subclassed
* Checks IDE Break Mode before dispatching (vba6.dll#EbMode).
* Set Compilation Argument DEBUG_MODE = 0, if compiling for release distribution

TODO:
* Add a tiny Call Stub to allow multiple subclasses per window

