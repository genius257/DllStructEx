#include "../au3pm/au3unit/Unit/assert.au3"
#include "../DllStructEx.au3"

$oA = DllStructExCreate("IDispatch *ref;")
$oB = DllStructExCreate("INT val;")

GLobal $tObject = DllStructCreate($__g_DllStructEx_tagObject, Ptr($oB) - 8)

assertEquals(1, $tObject.RefCount)

$oA.ref = $oB

assertEquals(2, $tObject.RefCount)

$oB = Null

assertEquals(1, $tObject.RefCount)

$oB = $oA.ref

assertEquals(2, $tObject.RefCount)

$oA.ref = $oB

assertEquals(2, $tObject.RefCount)

$oB.val = 123
$oB = Null

assertEquals(1, $tObject.RefCount)

assertEquals(123, $oA.ref.val)

assertEquals(1, $tObject.RefCount)

;$oA.ref = 0

;assertEquals(1, $tObject.RefCount)
