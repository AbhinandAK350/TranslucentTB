// Standard API
#include <chrono>
#include <mutex>
#include <sstream>
#include <string>
#include <thread>
#include <unordered_map>

// Windows API
#include "arch.h"
#include <PathCch.h>
#include <ShlObj.h>

// Local stuff
#include "appvisibilitysink.hpp"
#include "autofree.hpp"
#include "autostart.hpp"
#include "blacklist.hpp"
#include "common.hpp"
#include "config.hpp"
#include "createinstance.hpp"
#include "eventhook.hpp"
#include "messagewindow.hpp"
#include "resource.h"
#include "swcadata.hpp"
#include "traycontextmenu.hpp"
#include "ttberror.hpp"
#include "ttblog.hpp"
#include "util.hpp"
#ifdef STORE
#include "uwp.hpp"
#endif
#include "win32.hpp"
#include "window.hpp"
#include "windowclass.hpp"

#pragma region Data

HWND hwnd;

enum class EXITREASON {
	NewInstance,		// New instance told us to exit
	UserAction,			// Triggered by the user
	UserActionNoSave	// Triggered by the user, but doesn't saves config
};

static struct {
	EXITREASON exit_reason = EXITREASON::UserAction;
	std::mutex taskbars_mutex;
	Window main_taskbar;
	std::unordered_map<HMONITOR, std::pair<Window, const Config::TASKBAR_APPEARANCE *>> taskbars;
	bool should_show_peek = true;
	bool is_running = true;
	std::wstring config_folder;
	std::wstring config_file;
	std::wstring exclude_file;
	bool peek_active = false;
	bool start_opened = false;
} run;

#pragma endregion

#pragma region That one function that does all the magic

void SetWindowBlur(const Window &window, const swca::ACCENT &appearance, const uint32_t &color)
{
	if (user32::SetWindowCompositionAttribute)
	{
		static std::unordered_map<Window, bool> is_normal;

		swca::ACCENTPOLICY policy = {
			appearance,
			2,
			(color & 0xFF00FF00) + ((color & 0x00FF0000) >> 16) + ((color & 0x000000FF) << 16),
			0
		};

		if (policy.nAccentState == swca::ACCENT::ACCENT_NORMAL)
		{
			if (is_normal.count(window) == 0 || !is_normal[window])
			{
				// WM_THEMECHANGED makes the taskbar reload the theme and reapply the normal effect.
				// Gotta memoize it because constantly sending it makes explorer's CPU usage jump.
				window.send_message(WM_THEMECHANGED);
				is_normal[window] = true;
			}
			return;
		}
		else if (policy.nAccentState == swca::ACCENT::ACCENT_ENABLE_FLUENT && policy.nColor >> 24 == 0x00)
		{
			// Fluent mode doesn't likes a completely 0 opacity
			policy.nColor = (0x01 << 24) + (policy.nColor & 0x00FFFFFF);
		}

		swca::WINCOMPATTRDATA data = {
			swca::WindowCompositionAttribute::WCA_ACCENT_POLICY,
			&policy,
			sizeof(policy)
		};

		user32::SetWindowCompositionAttribute(window, &data);
		is_normal[window] = false;
	}
}

#pragma endregion

#pragma region Configuration

void GetPaths()
{
#ifndef STORE
	AutoFree::CoTaskMem<wchar_t> appDataSafe;
	ErrorHandle(SHGetKnownFolderPath(FOLDERID_RoamingAppData, KF_FLAG_DEFAULT, NULL, appDataSafe.put()), Error::Level::Fatal, L"Failed to determine configuration files locations!");
	const wchar_t *appData = appDataSafe.get();
#else
	try
	{
		winrt::hstring appData_str = UWP::GetApplicationFolderPath(UWP::FolderType::Roaming);
		const wchar_t *appData = appData_str.c_str();
#endif

	AutoFree::Local<wchar_t> configFolder;
	AutoFree::Local<wchar_t> configFile;
	AutoFree::Local<wchar_t> excludeFile;

	ErrorHandle(PathAllocCombine(appData, NAME, PATHCCH_ALLOW_LONG_PATHS, configFolder.put()), Error::Level::Fatal, L"Failed to combine AppData folder and application name!");
	ErrorHandle(PathAllocCombine(configFolder.get(), CONFIG_FILE, PATHCCH_ALLOW_LONG_PATHS, configFile.put()), Error::Level::Fatal, L"Failed to combine config folder and config file!");
	ErrorHandle(PathAllocCombine(configFolder.get(), EXCLUDE_FILE, PATHCCH_ALLOW_LONG_PATHS, excludeFile.put()), Error::Level::Fatal, L"Failed to combine config folder and exclude file!");

	run.config_folder = configFolder.get();
	run.config_file = configFile.get();
	run.exclude_file = excludeFile.get();

#ifdef STORE
	}
	catch (const winrt::hresult_error &error)
	{
		ErrorHandle(error.code(), Error::Level::Fatal, L"Getting application folder paths failed!");
	}
#endif
}

void ApplyStock(const std::wstring &filename)
{
	std::wstring exeFolder_str = win32::GetExeLocation();
	exeFolder_str.erase(exeFolder_str.find_last_of(LR"(/\)") + 1);

	AutoFree::Local<wchar_t> stockFile;
	if (!ErrorHandle(PathAllocCombine(exeFolder_str.c_str(), filename.c_str(), PATHCCH_ALLOW_LONG_PATHS, stockFile.put()), Error::Level::Error, L"Failed to combine executable folder and config file!"))
	{
		return;
	}

	AutoFree::Local<wchar_t> configFile;
	if (!ErrorHandle(PathAllocCombine(run.config_folder.c_str(), filename.c_str(), PATHCCH_ALLOW_LONG_PATHS, configFile.put()), Error::Level::Error, L"Failed to combine config folder and config file!"))
	{
		return;
	}

	if (!win32::IsDirectory(run.config_folder))
	{
		if (!CreateDirectory(run.config_folder.c_str(), NULL))
		{
			LastErrorHandle(Error::Level::Error, L"Creating configuration files directory failed!");
			return;
		}
	}

	if (!CopyFile(stockFile.get(), configFile.get(), FALSE))
	{
		LastErrorHandle(Error::Level::Error, L"Copying stock configuration file failed!");
	}
}

bool CheckAndRunWelcome()
{
	if (!win32::IsDirectory(run.config_folder))
	{
		// String concatenation is hard OK
		std::wostringstream message;
		message <<
			L"Welcome to " NAME L"!\n\n"
			L"You can tweak the taskbar's appearance with the tray icon. If it's your cup of tea, you can also edit the configuration files, located at \"" <<
			run.config_folder <<
			L"\"\n\nDo you agree to the GPLv3 license?";

		if (MessageBox(Window::NullWindow, message.str().c_str(), NAME, MB_ICONINFORMATION | MB_YESNO | MB_SETFOREGROUND) != IDYES)
		{
			return false;
		}
	}
	if (!win32::FileExists(run.config_file))
	{
		ApplyStock(CONFIG_FILE);
	}
	if (!win32::FileExists(run.exclude_file))
	{
		ApplyStock(EXCLUDE_FILE);
	}
	return true;
}

#pragma endregion

#pragma region Utilities

void RefreshHandles()
{
	if (Config::VERBOSE)
	{
		Log::OutputMessage(L"Refreshing taskbar handles.");
	}

	std::lock_guard guard(run.taskbars_mutex);

	// Older handles are invalid, so clear the map to be ready for new ones
	run.taskbars.clear();

	run.main_taskbar = Window::Find(L"Shell_TrayWnd");
	run.taskbars[run.main_taskbar.monitor()] = { run.main_taskbar, &Config::REGULAR_APPEARANCE };

	for (const Window secondtaskbar : Window::FindEnum(L"Shell_SecondaryTrayWnd"))
	{
		run.taskbars[secondtaskbar.monitor()] = { secondtaskbar, &Config::REGULAR_APPEARANCE };
	}
}

#pragma endregion

#pragma region Tray

void RefreshAutostartMenu(HMENU menu, const Autostart::StartupState &state)
{
	TrayContextMenu::RefreshBool(IDM_AUTOSTART, menu, !(state == Autostart::StartupState::DisabledByUser
		|| state == Autostart::StartupState::DisabledByPolicy
		|| state == Autostart::StartupState::EnabledByPolicy),
		TrayContextMenu::ControlsEnabled);

	TrayContextMenu::RefreshBool(IDM_AUTOSTART, menu, state == Autostart::StartupState::Enabled
		|| state == Autostart::StartupState::EnabledByPolicy,
		TrayContextMenu::Toggle);

	std::wstring autostart_text;
	switch (state)
	{
	case Autostart::StartupState::DisabledByUser:
		autostart_text = L"Startup has been disabled in Task Manager";
		break;
	case Autostart::StartupState::DisabledByPolicy:
		autostart_text = L"Startup has been disabled in Group Policy";
		break;
	case Autostart::StartupState::EnabledByPolicy:
		autostart_text = L"Startup has been enabled in Group Policy";
		break;
	case Autostart::StartupState::Enabled:
	case Autostart::StartupState::Disabled:
		autostart_text = L"Open at boot";
	}
	TrayContextMenu::ChangeItemText(menu, IDM_AUTOSTART, std::move(autostart_text));
}

void RefreshMenu(HMENU menu)
{
	TrayContextMenu::RefreshBool(IDM_AUTOSTART, menu, false, TrayContextMenu::ControlsEnabled);
	TrayContextMenu::RefreshBool(IDM_AUTOSTART, menu, false, TrayContextMenu::Toggle);
	TrayContextMenu::ChangeItemText(menu, IDM_AUTOSTART, L"Querying startup state...");
	Autostart::GetStartupState().Completed([menu](auto info, ...)
	{
		RefreshAutostartMenu(menu, info.GetResults());
	});
}

#pragma endregion

#pragma region Main logic

BOOL CALLBACK EnumWindowsProcess(const HWND hWnd, LPARAM)
{
	const Window window(hWnd);
	hwnd = hWnd;
	// DWMWA_CLOAKED should take care of checking if it's on the current desktop.
	// But that's undocumented behavior.
	// Do both but with on_current_desktop last.
	if (window.visible() && window.state() == SW_MAXIMIZE && !window.get_attribute<BOOL>(DWMWA_CLOAKED) &&
		!Blacklist::IsBlacklisted(window) && window.on_current_desktop() && run.taskbars.count(window.monitor()) != 0)
	{
		auto &taskbar = run.taskbars.at(window.monitor());
		if (Config::MAXIMISED_ENABLED)
		{
			taskbar.second = &Config::MAXIMISED_APPEARANCE;
		}

		if (Config::PEEK == Config::PEEK::Dynamic)
		{
			if (Config::PEEK_ONLY_MAIN)
			{
				if (taskbar.first == run.main_taskbar)
				{
					run.should_show_peek = true;
				}
			}
			else
			{
				run.should_show_peek = true;
			}
		}
	}
	return true;
}

void SetTaskbarBlur()
{
	const Window window(hwnd);
	static uint8_t counter = 10;

	std::lock_guard guard(run.taskbars_mutex);
	if (counter >= 10)	// Change this if you want to change the time it takes for the program to update.
	{					// 1 = Config::SLEEP_TIME; we use 10 (assuming the default configuration value of 10),
						// because the difference is less noticeable and it has no large impact on CPU.
						// We can change this if we feel that CPU is more important than response time.
		run.should_show_peek = (Config::PEEK == Config::PEEK::Enabled);

		for (auto &[_, pair] : run.taskbars)
		{
			pair.second = &Config::REGULAR_APPEARANCE; // Reset taskbar state
		}
		if (Config::MAXIMISED_ENABLED || Config::PEEK == Config::PEEK::Dynamic)
		{
			EnumWindows(&EnumWindowsProcess, NULL);
		}

		const Window fg_window = Window::ForegroundWindow();
		if (fg_window != Window::NullWindow && run.taskbars.count(fg_window.monitor()) != 0)
		{
			
			if (window.visible() && window.state() == SW_MAXIMIZE && !window.get_attribute<BOOL>(DWMWA_CLOAKED) &&
				!Blacklist::IsBlacklisted(window) && window.on_current_desktop() && run.taskbars.count(window.monitor()) != 0)
			{
				if (Config::CORTANA_ENABLED && !run.start_opened && !fg_window.get_attribute<BOOL>(DWMWA_CLOAKED))
				{
					const auto title = fg_window.filename();
					if (Util::IgnoreCaseStringEquals(*title, L"SearchUI.exe") || Util::IgnoreCaseStringEquals(*title, L"SearchApp.exe"))
					{
						run.taskbars.at(fg_window.monitor()).second = &Config::MAXIMISED_APPEARANCE;
					}
				}
				else
				{
					const auto title = fg_window.filename();
					if (Util::IgnoreCaseStringEquals(*title, L"SearchUI.exe") || Util::IgnoreCaseStringEquals(*title, L"SearchApp.exe"))
					{
						run.taskbars.at(fg_window.monitor()).second = &Config::CORTANA_APPEARANCE;
					}
				}

				if (Config::START_ENABLED && run.start_opened)
				{
					run.taskbars.at(fg_window.monitor()).second = &Config::MAXIMISED_APPEARANCE;
				}
				else
				{
					run.taskbars.at(fg_window.monitor()).second = &Config::START_APPEARANCE;
				}
			}
		}

		// Put this between Start/Cortana and Task view/Timeline
		// Task view and Timeline show over Aero Peek, but not Start or Cortana
		if (Config::MAXIMISED_ENABLED && Config::MAXIMISED_REGULAR_ON_PEEK && run.peek_active)
		{
			for (auto &[_, pair] : run.taskbars)
			{
				pair.second = &Config::REGULAR_APPEARANCE;
			}
		}

		if (fg_window != Window::NullWindow)
		{
			const static bool timeline_av = win32::IsAtLeastBuild(MIN_FLUENT_BUILD);
			if (Config::TIMELINE_ENABLED && (timeline_av
				? (*fg_window.classname() == CORE_WINDOW && Util::IgnoreCaseStringEquals(*fg_window.filename(), L"Explorer.exe"))
				: (*fg_window.classname() == L"MultitaskingViewFrame")))
			{
				for (auto &[_, pair] : run.taskbars)
				{
					pair.second = &Config::TIMELINE_APPEARANCE;
				}
			}
		}

		counter = 0;
	}
	else
	{
		counter++;
	}

	for (const auto &[_, pair] : run.taskbars)
	{
		const Config::TASKBAR_APPEARANCE &appearance = *pair.second;
		SetWindowBlur(pair.first, appearance.ACCENT, appearance.COLOR);
	}
}

#pragma endregion

#pragma region Startup

long ExitApp(const EXITREASON &reason, ...)
{
	run.exit_reason = reason;
	PostQuitMessage(0);
	return 0;
}

void InitializeTray(const HINSTANCE &hInstance)
{
	static MessageWindow window(L"TrayWindow", NAME, hInstance);

	window.RegisterCallback(NEW_TTB_INSTANCE, std::bind(&ExitApp, EXITREASON::NewInstance));

	window.RegisterCallback(WM_DISPLAYCHANGE, [](...)
	{
		RefreshHandles();
		return 0;
	});

	window.RegisterCallback(WM_TASKBARCREATED, [](...)
	{
		RefreshHandles();
		return 0;
	});

	window.RegisterCallback(WM_CLOSE, std::bind(&ExitApp, EXITREASON::UserAction));

	window.RegisterCallback(WM_QUERYENDSESSION, [](WPARAM, const LPARAM lParam)
	{
		if (lParam & ENDSESSION_CLOSEAPP)
		{
			// The app is being queried if it can close for an update.
			RegisterApplicationRestart(NULL, NULL);
		}
		return TRUE;
	});

	window.RegisterCallback(WM_ENDSESSION, [](const WPARAM wParam, const LPARAM lParam)
	{
		if (!(lParam & ENDSESSION_CLOSEAPP && !wParam))
		{
			// The app is being closed for an update or shutdown.
			Config::Save(run.config_file);
		}

		return 0;
	});


	if (!Config::NO_TRAY)
	{
		static TrayContextMenu tray(window, MAKEINTRESOURCE(TRAYICON), MAKEINTRESOURCE(IDR_POPUP_MENU), hInstance);

		tray.RegisterContextMenuCallback(IDM_AUTOSTART, []
		{
			Autostart::GetStartupState().Completed([](auto info, ...)
			{
				Autostart::SetStartupState(info.GetResults() == Autostart::StartupState::Enabled ? Autostart::StartupState::Disabled : Autostart::StartupState::Enabled);
			});
		});
		tray.RegisterContextMenuCallback(IDM_EXIT, std::bind(&ExitApp, EXITREASON::UserAction));

		tray.RegisterCustomRefresh(RefreshMenu);
	}
}

int WINAPI wWinMain(const HINSTANCE hInstance, HINSTANCE, wchar_t *, int)
{
	win32::HardenProcess();
	try
	{
		winrt::init_apartment(winrt::apartment_type::multi_threaded);
	}
	catch (const winrt::hresult_error &error)
	{
		ErrorHandle(error.code(), Error::Level::Fatal, L"Initialization of Windows Runtime failed.");
	}

	// If there already is another instance running, tell it to exit
	if (!win32::IsSingleInstance())
	{
		Window::Find(L"TrayWindow", NAME).send_message(NEW_TTB_INSTANCE);
	}

	// Get configuration file paths
	GetPaths();

	// If the configuration files don't exist, restore the files and show welcome to the users
	if (!CheckAndRunWelcome())
	{
		return EXIT_FAILURE;
	}

	// Parse our configuration
	Config::Parse(run.config_file);
	Blacklist::Parse(run.exclude_file);

	// Initialize GUI
	InitializeTray(hInstance);

	// Populate our map
	RefreshHandles();

	// Undoc'd, allows to detect when Aero Peek starts and stops
	EventHook peek_hook(
		0x21,
		0x22,
		[](const DWORD event, ...)
		{
			run.peek_active = event == 0x21;
		},
		WINEVENT_OUTOFCONTEXT
	);

	// Detect additional monitor connect/disconnect
	EventHook creation_hook(
		EVENT_OBJECT_CREATE,
		EVENT_OBJECT_DESTROY,
		[](DWORD, const Window &window, ...)
		{
			if (window.valid())
			{
				if (const auto classname = window.classname(); *classname == L"Shell_TrayWnd" || *classname == L"Shell_SecondaryTrayWnd")
				{
					RefreshHandles();
				}
			}
		},
		WINEVENT_OUTOFCONTEXT
	);

	// Register our start menu detection sink
	auto app_visibility = create_instance<IAppVisibility>(CLSID_AppVisibility);
	DWORD av_cookie = 0;
	if (app_visibility)
	{
		auto av_sink = winrt::make<AppVisibilitySink>(run.start_opened);
		ErrorHandle(app_visibility->Advise(av_sink.get(), &av_cookie), Error::Level::Log, L"Failed to register app visibility sink.");
	}

	std::thread swca_thread([]
	{
		try
		{
			winrt::init_apartment(winrt::apartment_type::single_threaded);
		}
		catch (const winrt::hresult_error &error)
		{
			ErrorHandle(error.code(), Error::Level::Fatal, L"Initialization of Windows Runtime failed.");
		}

		while (run.is_running)
		{
			SetTaskbarBlur();
			std::this_thread::sleep_for(std::chrono::milliseconds(Config::SLEEP_TIME));
		}
	});

	MSG msg;
	BOOL ret;
	while ((ret = GetMessage(&msg, NULL, 0, 0)) != 0)
	{
		if (ret != -1)
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		else
		{
			LastErrorHandle(Error::Level::Fatal, L"GetMessage failed!");
		}
	}

	run.is_running = false;
	swca_thread.join(); // Wait for our worker thread to exit.

	if (av_cookie)
	{
		ErrorHandle(app_visibility->Unadvise(av_cookie), Error::Level::Log, L"Failed to unregister app visibility sink.");
	}

	// If it's a new instance, don't save or restore taskbar to default
	if (run.exit_reason != EXITREASON::NewInstance)
	{
		if (run.exit_reason != EXITREASON::UserActionNoSave)
		{
			Config::Save(run.config_file);
		}

		// Restore default taskbar appearance
		for (const auto &taskbar : run.taskbars)
		{
			SetWindowBlur(taskbar.second.first, swca::ACCENT::ACCENT_NORMAL, NULL);
		}
	}

	return EXIT_SUCCESS;
}

#pragma endregion