using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MicroSipAutoConfig {
	class Program {
		private const string SOURCE_PATH = @"\\budzdorov.ru\NETLOGON\MicroSip";
		private const string CONTACTS_FILE = "Contacts.xml";
		private const string USER_LIST_FILE = "UserList.txt";
		private const string SETTINGS_FILE = "MicroSIP.ini";
		private const string DISTRIB_FOLDER = "Distrib";
		private const string DESTINATION_PATH = @"C:\Users\@userName\AppData\Local\MicroSIP\";

		static void Main(string[] args) {
			if (!Directory.Exists(@"C:\Temp\"))
				try {
					Directory.CreateDirectory(@"C:\Temp\");
				} catch (Exception) { }

			if (File.Exists(Logging.LOG_FILE_NAME))
				try {
					File.Delete(Logging.LOG_FILE_NAME);
				} catch (Exception) { }

			Logging.ToLog("Запуск");

			if (!Directory.Exists(SOURCE_PATH)) {
				Logging.ToLog("Не удается найти папку (получить доступ): " + SOURCE_PATH);
				return;
			}

			try {
				string currentUserSID = string.Empty;
				RegistryKey profileList = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList");
				string[] subKeys = profileList.GetSubKeyNames();
				foreach (string subKey in subKeys) {
					try {
						RegistryKey profile = profileList.OpenSubKey(subKey);
						string value = profile.GetValue("ProfileImagePath").ToString();

						if (value.EndsWith(Environment.UserName)) {
							currentUserSID = subKey;
							break;
						}
					} catch (Exception e) {
						Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					}
				}

				if (!string.IsNullOrEmpty(currentUserSID)) {
					string regPath = currentUserSID + "\\Software\\MicroSIP";
					Logging.ToLog("Удаление старых записей из реестра, ветка: " + regPath);
					Registry.Users.DeleteSubKeyTree(regPath);
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}






			string destinationPath = DESTINATION_PATH.Replace("@userName", Environment.UserName);
			string processToStart = Path.Combine(destinationPath, "microsip.exe");

			try {
				Process[] microSipProcesses = Process.GetProcessesByName("microsip");
				if (microSipProcesses.Length > 0) {
					foreach (Process process in microSipProcesses) {
						string query = "Select * From Win32_Process Where ProcessID = " + process.Id;
						string processOwner = string.Empty;
						ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
						ManagementObjectCollection processList = searcher.Get();

						foreach (ManagementObject obj in processList) {
							string[] argList = new string[] { string.Empty, string.Empty };
							int returnVal = Convert.ToInt32(obj.InvokeMethod("GetOwner", argList));
							if (returnVal == 0) 
								processOwner = argList[0];
						}

						if (!processOwner.Equals(Environment.UserName))
							continue;

						Logging.ToLog("Завершение запущенного процесса microsip.exe id: " + process.Id);
						process.Kill();
					}
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			if (!Directory.Exists(destinationPath)) {
				Logging.ToLog("Создание папки для установки: " + destinationPath);
				try {
					Directory.CreateDirectory(destinationPath);
				} catch (Exception e) {
					Logging.ToLog("Не удалось создать папку для установки, завершение");
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
					return;
				}
			}

			Logging.ToLog("Проверка установочных файлов MicroSIP в папке: " + destinationPath);
			try {
				string[] distribFiles = Directory.GetFiles(Path.Combine(SOURCE_PATH, DISTRIB_FOLDER));
				foreach (string distibFile in distribFiles) {
					string fileName = Path.GetFileName(distibFile);
					string destFile = Path.Combine(destinationPath, fileName);
					if (!File.Exists(destFile)) {
						Logging.ToLog("Копирование файла: " + destFile);
						File.Copy(distibFile, destFile);
						continue;
					}

					DateTime distibFileCreationTime = File.GetCreationTime(distibFile);
					DateTime destFileCreationTime = File.GetCreationTime(destFile);

					if (distibFileCreationTime.Equals(destFile)) {
						Logging.ToLog("Копирование файла: " + destFile);
						File.Copy(distibFile, destFile, true);
					}
				}
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			string sourceContactsFile = Path.Combine(SOURCE_PATH, CONTACTS_FILE);
			if (File.Exists(sourceContactsFile)) {
				string destContactsFile = Path.Combine(destinationPath, CONTACTS_FILE);
				Logging.ToLog("Копирование файла с контактами: " + destContactsFile);
				try {
					File.Copy(sourceContactsFile, destContactsFile, true);
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			} else {
				Logging.ToLog("Не удалось найти файл с контактами: " + sourceContactsFile);
			}

			string sourceSettingsFile = Path.Combine(SOURCE_PATH, SETTINGS_FILE);
			string destSettingsFile = Path.Combine(destinationPath, SETTINGS_FILE);
			if (!File.Exists(destSettingsFile)) {
				Logging.ToLog("Копирование файла с настройками: " + destSettingsFile);
				try {
					File.Copy(sourceSettingsFile, destSettingsFile);
				} catch (Exception e) {
					Logging.ToLog(e.StackTrace + Environment.NewLine + e.StackTrace);
				}
			}

			string label = string.Empty;
			string userPhone = string.Empty;
			string userListFile = Path.Combine(SOURCE_PATH, USER_LIST_FILE);
			if (File.Exists(userListFile)) {
				Logging.ToLog("Считывание списка имеющихся пользователей: " + userListFile);
				try {
					string[] usersUnfo = File.ReadAllLines(userListFile);
					string currentUser = Environment.UserName.ToLower();
					string currentUserInfo = string.Empty;

					foreach (string userInfo in usersUnfo) {
						if (!userInfo.ToLower().StartsWith(currentUser))
							continue;

						currentUserInfo = userInfo;
						break;
					}

					if (string.IsNullOrEmpty(currentUserInfo))
						Logging.ToLog("Не удалось найти параметры для текущего пользователя");
					else {
						string[] userInfoSplitted = currentUserInfo.Split('|');
						if (userInfoSplitted.Length != 3)
							Logging.ToLog("Неверный формат строки: " + currentUserInfo);
						else {
							Logging.ToLog("Параметры текущего пользователя: " + currentUserInfo);
							label = userInfoSplitted[1];
							userPhone = userInfoSplitted[2];
						}
					}
				} catch (Exception e) {
					Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
				}
			} else
				Logging.ToLog("Не удалось найти файл (получить доступ): " + userListFile);

			Logging.ToLog("Изменение файла настроек: " + destSettingsFile);
			try {
				List<string> settings = File.ReadAllLines(destSettingsFile).ToList();

				Dictionary<string, string> settingsToChange = new Dictionary<string, string> {
					{ "label=", label },
					{ "server=", "172.16.210.140" },
					{ "domain=", "172.16.210.140" },
					{ "proxy=", "" },
					{ "username=", userPhone },
					{ "password=", "602b7dadeea558e0ad931683e5c35613" },
					{ "publish=", "1" },
					{ "audioCodecs=", "PCMA/8000/1 PCMU/8000/1" },
					{ "accountId=", "1" },
					{ "recordingPath=", @"C:\Users\" + Environment.UserName + @"\Desktop\Recordings" },
					{ "[Account", "1]" },
					{ "authID=", userPhone },
					{ "displayName=", label }
				};

				foreach (string settingsLine in settings.ToArray()) {
					foreach (KeyValuePair<string, string> settingToChange in settingsToChange) {
						if (!settingsLine.StartsWith(settingToChange.Key))
							continue;

						Logging.ToLog("Параметр '" + settingToChange.Key + "' новое значение: " + settingToChange.Value);
						settings[settings.IndexOf(settingsLine)] = settingToChange.Key + settingToChange.Value;
						break;
					}
				}

				Logging.ToLog("Запись настроек в файл: " + destinationPath);
				File.WriteAllLines(destSettingsFile, settings);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			Logging.ToLog("Создание ярлыка на рабочем столе для MicroSIP.exe");
			try {

				object shDesktop = (object)"Desktop";
				IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
				string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + @"\MicroSIP.lnk";
				IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcutAddress);
				shortcut.Description = "Ярлык для MicroSIP";
				shortcut.TargetPath = processToStart;
				shortcut.Save();
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}

			Logging.ToLog("Запуск MicroSip.exe: " + processToStart);
			try {
				Process.Start(processToStart);
			} catch (Exception e) {
				Logging.ToLog(e.Message + Environment.NewLine + e.StackTrace);
			}
		}
	}
}
