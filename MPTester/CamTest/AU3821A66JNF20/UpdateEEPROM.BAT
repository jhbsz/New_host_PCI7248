CLS
@ECHO OFF

call %1

if not %errorlevel% == 0 goto Err
copy %2\UpdateResult\UpdatePASS %2\UpdatePASS
	
@ECHO ------------------------: "Update successfully."

goto End

	:Err
	copy %2\UpdateResult\UpdateFAIL %2\UpdateFAIL	
	@ECHO ------------------------: "Update error."

:End

