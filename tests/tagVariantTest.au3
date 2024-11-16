#include "../au3pm/au3unit/Unit/assert.au3"
#include "../DllStructEx.au3"
#include "../DllStructEx.debug.au3"

$tagVARIANT = _
"  union {"& _
"    struct {"& _
"      VARTYPE vt;"& _
"      WORD    wReserved1;"& _
"      WORD    wReserved2;"& _
"      WORD    wReserved3;"& _
"      union {"& _
"        LONGLONG     llVal;"& _
"        LONG         lVal;"& _
"        BYTE         bVal;"& _
"        SHORT        iVal;"& _
"        FLOAT        fltVal;"& _
"        DOUBLE       dblVal;"& _
"        VARIANT_BOOL boolVal;"& _
"        VARIANT_BOOL __OBSOLETE__VARIANT_BOOL;"& _
"        SCODE        scode;"& _
"        CY           cyVal;"& _
"        DATE         date;"& _
"        BSTR         bstrVal;"& _
"        IUnknown     *punkVal;"& _
"        IDispatch    *pdispVal;"& _
"        SAFEARRAY    *parray;"& _
"        BYTE         *pbVal;"& _
"        SHORT        *piVal;"& _
"        LONG         *plVal;"& _
"        LONGLONG     *pllVal;"& _
"        FLOAT        *pfltVal;"& _
"        DOUBLE       *pdblVal;"& _
"        VARIANT_BOOL *pboolVal;"& _
"        VARIANT_BOOL *__OBSOLETE__VARIANT_PBOOL;"& _
"        SCODE        *pscode;"& _
"        CY           *pcyVal;"& _
"        DATE         *pdate;"& _
"        BSTR         *pbstrVal;"& _
"        IUnknown     **ppunkVal;"& _
"        IDispatch    **ppdispVal;"& _
"        SAFEARRAY    **pparray;"& _
"        VARIANT      *pvarVal;"& _
"        PVOID        byref;"& _
"        CHAR         cVal;"& _
"        USHORT       uiVal;"& _
"        ULONG        ulVal;"& _
"        ULONGLONG    ullVal;"& _
"        INT          intVal;"& _
"        UINT         uintVal;"& _
"        DECIMAL      *pdecVal;"& _
"        CHAR         *pcVal;"& _
"        USHORT       *puiVal;"& _
"        ULONG        *pulVal;"& _
"        ULONGLONG    *pullVal;"& _
"        INT          *pintVal;"& _
"        UINT         *puintVal;"& _
"        struct {"& _
"          PVOID       pvRecord;"& _
"          IRecordInfo *pRecInfo;"& _
"        } __VARIANT_NAME_4;"& _
"      } __VARIANT_NAME_3;"& _
"    } __VARIANT_NAME_2;"& _
"    DECIMAL decVal;"& _
"  } __VARIANT_NAME_1;"

$tagDECIMAL = _
  "USHORT wReserved;"& _
  "union {"& _
  "  struct {"& _
  "    BYTE scale;"& _
  "    BYTE sign;"& _
  "  } DUMMYSTRUCTNAME;"& _
  "  USHORT signscale;"& _
  "} DUMMYUNIONNAME;"& _
  "ULONG  Hi32;"& _
  "union {"& _
  "  struct {"& _
  "    ULONG Lo32;"& _
  "    ULONG Mid32;"& _
  "  } DUMMYSTRUCTNAME2;"& _
  "  ULONGLONG Lo64;"& _
  "} DUMMYUNIONNAME2;"

$typeULONGLONG = "UINT64"
$typeVARTYPE = "USHORT"
$typeLONGLONG = "INT64"
$typeVARIANT_BOOL = "SHORT"
$typeSCODE = "LONG"
$typeDATE = "DOUBLE"
$typeBSTR = "PTR"
$typePVOID = "HANDLE"
$tagCY = "ULONG Lo;LONG Hi;"

$txVariant = DllStructExCreate($tagVARIANT)
assertEquals(0, @error)

assertEquals("BYTE __VARIANT_NAME_1[16];", DllStructExGetTranspiledStructString($txVariant))
