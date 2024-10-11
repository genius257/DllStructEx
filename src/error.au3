#include-once

Global $__DllStructEx_bErrorMessages = Not @Compiled

Func __DllStructEx_ConsoleWriteErrorLine($s, $line = @ScriptLineNumber)
    ConsoleWriteError($s&@CRLF)
    ConsoleWriteError(@ScriptFullPath&":"&$line&@CRLF)
EndFunc
