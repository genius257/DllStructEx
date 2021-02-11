# DllStructEx
 Extended DllStruct for AutoIt3

This project adds missing functionality to DLLStruct's:
* Being able to get the inital struct string
* Support for union's
* Allow nested structs

## drawbacks
Speed.

The soltion is implemented via a IDispatch interface. As a result, the calls to the object is slower than normal function calls.
