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

#region struct
$tx = DllStructExCreate("struct {INT a; INT b[2];} a[5];")
assertEquals(0, @error)
assertEquals("STRUCT;INT a;INT[2];INT;INT[2];INT;INT[2];INT;INT[2];INT;INT[2];ENDSTRUCT;", DllStructExGetTranspiledStructString($tx))
$tx.a.a = 123
assertEquals(123, DllStructGetData(DllStructExGetStruct($tx.a), 1))
assertEquals(123, $tx.a.a)
$tx.a(1).a = 1
$tx.a(2).a = 2
$tx.a(3).a = 3
$tx.a(4).a = 4
$tx.a(5).a = 5
assertEquals(1, DllStructGetData(DllStructExGetStruct($tx.a(1)), 1))
assertEquals(1, $tx.a(1).a)
assertEquals(2, DllStructGetData(DllStructExGetStruct($tx.a(2)), 1))
assertEquals(2, $tx.a(2).a)
assertEquals(3, DllStructGetData(DllStructExGetStruct($tx.a(3)), 1))
assertEquals(3, $tx.a(3).a)
assertEquals(4, DllStructGetData(DllStructExGetStruct($tx.a(4)), 1))
assertEquals(4, $tx.a(4).a)
assertEquals(5, DllStructGetData(DllStructExGetStruct($tx.a(5)), 1))
assertEquals(5, $tx.a(5).a)
#endregion

#region union
$tx = DllStructExCreate("union {INT a; INT b[2];} a[5];")
assertEquals(0, @error)
assertEquals("BYTE a[40];", DllStructExGetTranspiledStructString($tx))

$tx.a.a = 123
assertEquals(123, DllStructGetData(DllStructExGetStruct($tx.a), 1))
assertEquals(123, $tx.a.a)

$tx.a(1).a = 1
$tx.a(2).a = 2
$tx.a(3).a = 3
$tx.a(4).a = 4
$tx.a(5).a = 5
assertEquals(1, DllStructGetData(DllStructExGetStruct($tx.a(1)), 1))
assertEquals(1, $tx.a(1).a)
assertEquals(2, DllStructGetData(DllStructExGetStruct($tx.a(2)), 1))
assertEquals(2, $tx.a(2).a)
assertEquals(3, DllStructGetData(DllStructExGetStruct($tx.a(3)), 1))
assertEquals(3, $tx.a(3).a)
assertEquals(4, DllStructGetData(DllStructExGetStruct($tx.a(4)), 1))
assertEquals(4, $tx.a(4).a)
assertEquals(5, DllStructGetData(DllStructExGetStruct($tx.a(5)), 1))
assertEquals(5, $tx.a(5).a)
#endregion

#region pointer
    $tx = DllStructExCreate("INT *a[5];")
    assertEquals(0, @error)
    assertEquals("PTR a[5];", DllStructExGetTranspiledStructString($tx))

    #region setup test data
        $t = DllStructCreate("INT v[5];")
        DllStructSetData($t, 1, 1, 1)
        DllStructSetData($t, 1, 2, 2)
        DllStructSetData($t, 1, 3, 3)
        DllStructSetData($t, 1, 4, 4)
        DllStructSetData($t, 1, 5, 5)
        DllStructSetData(DllStructExGetStruct($tx), 1, DllStructGetPtr($t), 1)
        DllStructSetData(DllStructExGetStruct($tx), 1, DllStructGetPtr($t)+(@AutoItX64 ? 8 : 4), 2)
        DllStructSetData(DllStructExGetStruct($tx), 1, DllStructGetPtr($t)+(@AutoItX64 ? 16 : 8), 3)
        DllStructSetData(DllStructExGetStruct($tx), 1, DllStructGetPtr($t)+(@AutoItX64 ? 24 : 12), 4)
        DllStructSetData(DllStructExGetStruct($tx), 1, DllStructGetPtr($t)+(@AutoItX64 ? 32 : 16), 5)
    #endregion

    assertEquals(1, $tx.a(1))
    assertEquals(2, $tx.a(2))
    assertEquals(3, $tx.a(3))
    assertEquals(4, $tx.a(4))
    assertEquals(5, $tx.a(5))

    $tx.a(1) = 11
    assertEquals(0, @error)
    $tx.a(2) = 22
    assertEquals(0, @error)
    $tx.a(3) = 33
    assertEquals(0, @error)
    $tx.a(4) = 44
    assertEquals(0, @error)
    $tx.a(5) = 55
    assertEquals(0, @error)

    assertEquals(11, $tx.a(1))
    assertEquals(22, $tx.a(2))
    assertEquals(33, $tx.a(3))
    assertEquals(44, $tx.a(4))
    assertEquals(55, $tx.a(5))
#endregion

