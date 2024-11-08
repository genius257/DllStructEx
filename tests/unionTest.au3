#include "../au3pm/au3unit/Unit/assert.au3"
#include "../DllStructEx.au3"
#include "../DllStructEx.debug.au3"

$o = DllStructExCreate("union{INT a;INT b;INT c;} u;")
assertEquals(0, @error)
assertSame($o.u.a, $o.u.b)
assertSame($o.u.b, $o.u.c)
$o.u.a = 123
assertSame(123, $o.u.a)
assertSame(123, $o.u.b)
assertSame(123, $o.u.c)

$tagUnionA = "DWORD a;union{A a;B b;C c;union{INT d;} d;} b;"
$tagA = "INT a;"
$tagB = "INT b;"
$tagC = "INT c;"
$oUnionA = DllStructExCreate($tagUnionA)
assertEquals(0, @error)

;DllStructExDisplay($oUnionA)
$oA = $oUnionA.b.a
$oB = $oUnionA.b.b
$oC = $oUnionA.b.c
$oD = $oUnionA.b.d

assertSame(DllStructExGetPtr($oA), DllStructExGetPtr($oB))
assertSame(DllStructExGetPtr($oB), DllStructExGetPtr($oC))
assertSame(DllStructExGetPtr($oC), DllStructExGetPtr($oD))

$oA.a = 123
assertSame(123, $oA.a)
assertSame(123, $oB.b)
assertSame(123, $oC.c)
assertSame(123, $oD.d)
