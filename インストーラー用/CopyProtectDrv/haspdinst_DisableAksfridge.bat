ECHO OFF
CLS

ECHO HASP aksfridge.sys を無効化します.
ECHO.                       
PAUSE

IF NOT EXIST %SystemRoot%\system32\drivers\aksfridge.sys goto NOTHING

MOVE /Y %SystemRoot%\system32\drivers\aksfridge.sys old-aksfridge.old
IF NOT EXIST %SystemRoot%\system32\drivers\aksfridge.sys goto DONE

ECHO.
ECHO エラー：%SystemRoot%\system32\drivers\aksfridge.sys の無効化に失敗.
                       
GOTO END


:DONE
                       
ECHO.
ECHO HASP aksfridge.sys を無効化しました.
                       
GOTO END


:NOTHING
                       
ECHO.
ECHO HASP aksfridge.sys がシステム内に見つかりませんでした.
                       
GOTO END


:END

ECHO.
PAUSE
