# offload-tax-app
AI agent system for tax practice management

## Run it on your own computer

1. Put your Anthropic key in a file named `.env` in this folder: `ANTHROPIC_API_KEY=sk-ant-...`
2. Double-click `start-offload.bat`. It starts Offload at http://localhost:5000 (this computer only) and opens your browser.
3. First time: create an account on the sign-in screen. Close the black window to stop Offload.

Everything is saved as files in this folder - to-dos, calendar and knowledge base, plus your tasks, projects, chats and lessons (in the `store` folder) - so a synced folder like OneDrive keeps all of it. Settings > Data can back it up or move it to another computer.

## Run it as a desktop program (its own window)

Right-click `make-desktop-shortcut.ps1` > Run with PowerShell (once). It puts an **Offload** shortcut on your Desktop. Double-click it: Offload opens in its own window (no browser tabs, no black console window), and closing the window stops it. It needs Google Chrome or Microsoft Edge installed. Problems are written to `offload-desktop.log`.
