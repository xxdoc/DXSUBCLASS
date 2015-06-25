# DXSUBCLASS
Robust Subclassing DLL for debugging in the VB6 IDE

* Adapted from Karl E. Peterson --> http://vb.mvps.org/samples/HookXP/
* No Assembly, Dispatching Stubs, Thunks or Self Modifying code
* All Visual Basic (64 lines)
* Uses Windows subclassing API comctl32.dll#SetWindowSubclass
* Multiple Subclasses per hWnd
* no need to manage the subclassing chain, unsubclass in any order
* DEBUG_MODE checks VB6 IDE Break Mode before dispatching (vba6.dll#EbMode).
* Set Compilation Argument DEBUG_MODE = 0, if compiling for release distribution
* See CMinMax.cls for usage
