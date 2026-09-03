-- Quelle für "Watcher starten.app" (macOS-Doppelklick-Launcher, kein sichtbares Terminal).
-- Neu kompilieren nach Änderungen: osacompile -o "Watcher starten.app" start.applescript

set appPosixPath to POSIX path of (path to me)
set oldDelims to AppleScript's text item delimiters
set AppleScript's text item delimiters to "/"
set pathParts to text items of appPosixPath
-- appPosixPath endet auf ".../Watcher starten.app/" -> letzte zwei Teile (Appname, leerer Rest) entfernen
set pathParts to items 1 thru -3 of pathParts
set repoDir to (pathParts as text) & "/"
set AppleScript's text item delimiters to oldDelims

set pyBin to repoDir & ".venv/bin/python"

try
	do shell script "cd " & quoted form of repoDir & " && " & quoted form of pyBin & " start.py"
on error errMsg
	display alert "budget-watcher: Start fehlgeschlagen" message errMsg
end try
