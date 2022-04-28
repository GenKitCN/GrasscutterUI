for /F "tokens=2" %%s in ('certutil -dump %UserProfile%\.mitmproxy\mitmproxy-ca-cert.cer ^| findstr ^"^sha1^"') do (
		set SERIAL=%%s
	)
certutil -delstore root %SERIAL% >nul 2>nul