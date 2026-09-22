# Restore WMIC For Windows 11 (24h2/25H2)

Getting 'wmic' is not recognized as an internal or external command, operable program or batch file? Windows 11 24H2/25H2 updates removed WMIC. Pick one fix.

1. Download this repository.
2. Double-click "Restore WMIC via PowerShell.bat" and click Yes.

The installer is digitally signed.

Wait for SUCCESS - WMIC is working.

Check that WMIC works:

    wmic cpu get name
