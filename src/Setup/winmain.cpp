// Setup.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "Setup.h"
#include "FxHelper.h"
#include "UpdateRunner.h"
#include "MachineInstaller.h"
#include <cstdio>
#include <cstdarg>

CAppModule _Module;

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                      _In_opt_ HINSTANCE hPrevInstance,
                      _In_ LPWSTR lpCmdLine,
                      _In_ int nCmdShow)
{
	// Attempt to mitigate http://textslashplain.com/2015/12/18/dll-hijacking-just-wont-die
	SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32);

	int exitCode = -1;
	CString cmdLine(lpCmdLine);

    LogMessage(false, L"Start up installer: %s", lpCmdLine);

	if (cmdLine.Find(L"--checkInstall") >= 0) {
		// If we're already installed, exit as fast as possible
		if (!MachineInstaller::ShouldSilentInstall()) {
            LogMessage(false, L"Already installed");
			exitCode = 0;
			goto out;
		}

		// Make sure update.exe gets silent
		wcscat(lpCmdLine, L" --silent");
	}

	HRESULT hr = ::CoInitialize(NULL);
	ATLASSERT(SUCCEEDED(hr));

	AtlInitCommonControls(ICC_COOL_CLASSES | ICC_BAR_CLASSES);
	hr = _Module.Init(NULL, hInstance);

	bool isQuiet = (cmdLine.Find(L"-s") >= 0);
	bool weAreUACElevated = CUpdateRunner::AreWeUACElevated() == S_OK;
	bool explicitMachineInstall = (cmdLine.Find(L"--machine") >= 0);

	if (explicitMachineInstall || weAreUACElevated) {
        LogMessage(false, L"Want machine install");

		exitCode = MachineInstaller::PerformMachineInstallSetup();
		if (exitCode != 0) goto out;
		isQuiet = true;

		// Make sure update.exe gets silent
		if (explicitMachineInstall) {
			wcscat(lpCmdLine, L" --silent");
            LogMessage(false, L"Machine-wide installation was successful! Users will see the app once they log out / log in again.");
		}
    }
    else {
        LogMessage(false, L"Want standard install");
    }

	if (!CFxHelper::CanInstallDotNet4_5()) {
		// Explain this as nicely as possible and give up.
		MessageBox(0L, L"This program cannot run on Windows XP or before; it requires a later version of Windows.", L"Incompatible Operating System", 0);
		exitCode = E_FAIL;
		goto out;
	}

	if (!CFxHelper::IsDotNet45OrHigherInstalled()) {
		hr = CFxHelper::InstallDotNetFramework(isQuiet);
		if (FAILED(hr)) {
			exitCode = hr; // #yolo
			CUpdateRunner::DisplayErrorMessage(CString(L"Failed to install the .NET Framework, try installing .NET 4.5 or higher manually"), NULL);
			goto out;
		}
	
		// S_FALSE isn't failure, but we still shouldn't try to install
		if (hr != S_OK) {
			exitCode = 0;
			goto out;
		}
	}

	// If we're UAC-elevated, we shouldn't be because it will give us permissions
	// problems later. Just silently rerun ourselves.
	if (weAreUACElevated) {
		wchar_t buf[4096];
		HMODULE hMod = GetModuleHandle(NULL);
		GetModuleFileNameW(hMod, buf, 4096);

        LogMessage(false, L"we are UAC elevated, so restart %s, %s\n", buf, lpCmdLine);

		CUpdateRunner::ShellExecuteFromExplorer(buf, lpCmdLine);
		exitCode = 0;
		goto out;
	}

	exitCode = CUpdateRunner::ExtractUpdaterAndRun(lpCmdLine, false);

out:
	_Module.Term();
	::CoUninitialize();
	return exitCode;
}

void LogMessage(bool showMessageBox, const wchar_t* fmt, ...)
{
    wchar_t buff[1024];
    va_list args;
    va_start(args, fmt);
    _vsnwprintf(buff, sizeof(buff), fmt, args);
    va_end(args);
    OutputDebugString(buff);
    const char* tempDir = getenv("TEMP");
    if (tempDir) {
        char path[MAX_PATH];
        sprintf(path, "%s\\SquirrelSetup.log", tempDir);
        FILE* f = fopen(path, "a");
        if (f) {
            char cbuff[2048];
            size_t len = wcstombs(cbuff, buff, sizeof(cbuff)-2);
            cbuff[len++] = '\n';
            cbuff[len] = 0;
            fwrite(cbuff, 1, len, f);
            fclose(f);
        }
    }
    if (showMessageBox) {
        MessageBox(NULL, buff, L"Installer", MB_OK);
    }
}
