@echo off
set back=%cd%
for /R %%i in (.) do (
	cd "%%i"
		for %%q in (*.*) do (
				if "%%~xq" EQU ".json" ( 
						if "%%q" NEQ "Collection.postman_collection.json" ( 
							newman run "%back%\Collection.postman_collection.json" -d "%%q" --insecure > "output-report.log"
							@echo off
						)
				) 
		)
) 
cd %back%
	
