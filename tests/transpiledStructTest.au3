#include "../au3pm/au3unit/Unit/assert.au3"
#include "../DllStructEx.au3"

$oStruct = DllStructExCreate("WORD a;union {INT a;PTR b;BOOL c;} x;WORD b;")
assertEquals("WORD a;BYTE[2];BYTE x[4];WORD b;", DllStructExGetTranspiledStructString($oStruct))
assertEquals("INT a;PTR b;BOOL c;", DllStructExGetTranspiledStructString($oStruct.x))
$oStruct = DllStructExCreate("WORD a;struct {INT a;PTR b;BOOL c;} x;WORD b;")
assertEquals("WORD a;STRUCT;INT x;PTR;BOOL;ENDSTRUCT;WORD b;", DllStructExGetTranspiledStructString($oStruct))
assertEquals("INT a;PTR b;BOOL c;", DllStructExGetTranspiledStructString($oStruct.x))
$tagSub = "union {INT a;PTR b;BOOL c;} x;"
$oStruct = DllStructExCreate("WORD a;Sub x;WORD b;")
assertEquals("WORD a;STRUCT;BYTE x[4];ENDSTRUCT;WORD b;", DllStructExGetTranspiledStructString($oStruct))
assertEquals("BYTE x[4];", DllStructExGetTranspiledStructString($oStruct.x))
assertEquals("INT a;PTR b;BOOL c;", DllStructExGetTranspiledStructString($oStruct.x.x))
$tagSub = "INT a;PTR b;BOOL c;"
$oStruct = DllStructExCreate("WORD a;Sub x;WORD b;")
assertEquals("WORD a;STRUCT;INT x;PTR;BOOL;ENDSTRUCT;WORD b;", DllStructExGetTranspiledStructString($oStruct))
assertEquals("INT a;PTR b;BOOL c;", DllStructExGetTranspiledStructString($oStruct.x))
