#include "../au3pm/au3unit/Unit/assert.au3"
#include "../DllStructEx.au3"

$tagSomething = "  INT a;  "
$o = DllStructExCreate($tagSomething)
assertEquals(0, @error)
