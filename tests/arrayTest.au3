#include "../au3pm/au3unit/Unit/assert.au3"
#include "../DllStructEx.au3"

#region simple declaration
$tx = DllStructExCreate("INT a[5];")
assertEquals(0, @error)
assertEquals("INT a[5];", DllStructExGetTranspiledStructString($tx))
$t = DllStructExGetStruct($tx)
assertEquals(0, @error)
$tx.a(1) = 1
$tx.a(2) = 2
$tx.a(3) = 3
$tx.a(4) = 4
$tx.a(5) = 5
assertEquals(1, DllStructGetData($t, 1, 1))
assertEquals(1, $tx.a(1))
assertEquals(2, DllStructGetData($t, 1, 2))
assertEquals(2, $tx.a(2))
assertEquals(3, DllStructGetData($t, 1, 3))
assertEquals(3, $tx.a(3))
assertEquals(4, DllStructGetData($t, 1, 4))
assertEquals(4, $tx.a(4))
assertEquals(5, DllStructGetData($t, 1, 5))
assertEquals(5, $tx.a(5))
#endregion

