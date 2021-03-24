# DllStructEx
 Extended DllStruct for AutoIt3

This project adds missing functionality to DLLStruct's:
* Being able to get the inital struct string
* Support for union's
* Allow nested structs

Basically brining C-style structs to AutoIt3.

See the example for usage.

## drawbacks
Speed.

The solution is implemented via a IDispatch interface. As a result, the calls to the object is slower than normal function calls.
