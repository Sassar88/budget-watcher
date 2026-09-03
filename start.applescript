-- Quelle für "Watcher starten.app" (macOS-Doppelklick-Launcher, kein sichtbares Terminal).
-- Neu kompilieren nach Änderungen: osacompile -o "Watcher starten.app" start.applescript

set repoDir to POSIX path of (container of (path to me) as alias)
set pyBin to repoDir & ".venv/bin/python"

try
	do shell script "cd " & quoted form of repoDir & " && " & quoted form of pyBin & " start.py"
on error errMsg
	display alert "budget-watcher: Start fehlgeschlagen" message errMsg
end try
