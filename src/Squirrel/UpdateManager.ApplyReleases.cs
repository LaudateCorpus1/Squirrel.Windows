﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NuGet;
using Splat;
using System.Threading;
using Squirrel.Shell;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Core;

namespace Squirrel
{
    public sealed partial class UpdateManager
    {
        internal class ApplyReleasesImpl : IEnableLogger
        {
            readonly string rootAppDirectory;

            public ApplyReleasesImpl(string rootAppDirectory)
            {
                this.rootAppDirectory = rootAppDirectory;
            }

            public async Task<string> ApplyReleases(UpdateInfo updateInfo, bool silentInstall, bool attemptingFullInstall, Action<int> progress = null)
            {
                progress = progress ?? (_ => { });

                var release = await createFullPackagesFromDeltas(updateInfo.ReleasesToApply, updateInfo.CurrentlyInstalledVersion);
                progress(10);

                if (release == null) {
                    if (attemptingFullInstall) {
                        this.Log().Info("No release to install, running the app");
                        await invokePostInstall(updateInfo.CurrentlyInstalledVersion.Version, false, true, silentInstall);
                    }

                    progress(100);
                    return getDirectoryForRelease(updateInfo.CurrentlyInstalledVersion.Version).FullName;
                }

                var ret = await this.ErrorIfThrows(() => installPackageToAppDir(updateInfo, release), 
                    "Failed to install package to app dir");
                progress(30);

                var currentReleases = await this.ErrorIfThrows(() => updateLocalReleasesFile(),
                    "Failed to update local releases file");
                progress(50);

                var newVersion = currentReleases.MaxBy(x => x.Version).First().Version;
                executeSelfUpdate(newVersion);

                await this.ErrorIfThrows(() => invokePostInstall(newVersion, attemptingFullInstall, false, silentInstall),
                    "Failed to invoke post-install");
                progress(75);

                this.Log().Info("Starting fixPinnedExecutables");
                this.ErrorIfThrows(() => fixPinnedExecutables(updateInfo.FutureReleaseEntry.Version));

                this.Log().Info("Fixing up tray icons");

                var trayFixer = new TrayStateChanger();
                var appDir = new DirectoryInfo(Utility.AppDirForRelease(rootAppDirectory, updateInfo.FutureReleaseEntry));
                var allExes = appDir.GetFiles("*.exe").Select(x => x.Name).ToList();

                this.ErrorIfThrows(() => trayFixer.RemoveDeadEntries(allExes, rootAppDirectory, updateInfo.FutureReleaseEntry.Version.ToString()));
                progress(80);

                try {
                    var currentVersion = updateInfo.CurrentlyInstalledVersion != null ?
                        updateInfo.CurrentlyInstalledVersion.Version : null;

                    await cleanDeadVersions(currentVersion, newVersion);
                } catch (Exception ex) {
                    this.Log().WarnException("Failed to clean dead versions, continuing anyways", ex);
                }
                progress(100);

                return ret;
            }

            public async Task FullUninstall()
            {
                var currentRelease = getReleases().MaxBy(x => x.Name.ToVersion()).FirstOrDefault();

                this.Log().Info("Starting full uninstall");
                if (currentRelease.Exists) {
                    var version = currentRelease.Name.ToVersion();

                    try {
                        var squirrelAwareApps = SquirrelAwareExecutableDetector.GetAllSquirrelAwareApps(currentRelease.FullName);

                        if (isAppFolderDead(currentRelease.FullName)) throw new Exception("App folder is dead, but we're trying to uninstall it?");

                        var allApps = currentRelease.EnumerateFiles()
                            .Where(x => x.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                            .Where(x => !x.Name.StartsWith("squirrel.", StringComparison.OrdinalIgnoreCase) && !x.Name.StartsWith("update.", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        if (squirrelAwareApps.Count > 0) {
                            await squirrelAwareApps.ForEachAsync(async exe => {
                                using (var cts = new CancellationTokenSource()) { 
                                    cts.CancelAfter(10 * 1000);

                                    try {
                                        await Utility.InvokeProcessAsync(exe, String.Format("--squirrel-uninstall {0}", version), cts.Token);
                                    } catch (Exception ex) {
                                        this.Log().ErrorException("Failed to run cleanup hook, continuing: " + exe, ex);
                                    }
                                }
                            }, 1 /*at a time*/);
                        } else {
                            allApps.ForEach(x => RemoveShortcutsForExecutable(x.Name, ShortcutLocations.Defaults));
                        }

                        // NB: Some people attempt to uninstall apps while 
                        // they're still running. I cannot even.
                        var toKill = allApps
                            .SelectMany(x => Process.GetProcessesByName(x.Name.Replace(".exe", "")))
                            .ToList();

                        if (toKill.Count > 0) {
                            toKill.ForEach(x => x.Kill());
                            Thread.Sleep(750);
                        }
                    } catch (Exception ex) {
                        this.Log().WarnException("Failed to run pre-uninstall hooks, uninstalling anyways", ex);
                    }
                }

                fixPinnedExecutables(new Version(255, 255, 255, 255));

                bool didSucceedDeleting = false;
                const int retryAttempts = 10;
                for (int i = 0; i < retryAttempts; ++i) {
                    try {
                        await Utility.DeleteDirectory(rootAppDirectory);
                        didSucceedDeleting = true;
                    } catch (Exception) {
                        Thread.Sleep(1000); // Give the OS a second to release handles and we'll try again
                    }
                }

                if (!didSucceedDeleting) {
                    await this.ErrorIfThrows(() => Utility.DeleteDirectoryWithFallbackToNextReboot(rootAppDirectory),
                        "Failed to delete app directory: " + rootAppDirectory);
                }

                // NB: We drop this file here so that --checkInstall will ignore 
                // this folder - if we don't do this, users who "accidentally" run as 
                // administrator will find the app reinstalling itself on every
                // reboot
                File.WriteAllText(Path.Combine(rootAppDirectory, ".dead"), " ");
            }

            public Dictionary<ShortcutLocation, ShellLink> GetShortcutsForExecutable(string exeName, ShortcutLocation locations, string programArguments)
            {
                this.Log().Info("About to create shortcuts for {0}, rootAppDir {1}", exeName, rootAppDirectory);

                var releases = Utility.LoadLocalReleases(Utility.LocalReleaseFileForAppDir(rootAppDirectory));
                var thisRelease = Utility.FindCurrentVersion(releases);
                var updateExe = Path.Combine(rootAppDirectory, "update.exe");

                var zf = new ZipPackage(Path.Combine(
                    Utility.PackageDirectoryForAppDir(rootAppDirectory),
                    thisRelease.Filename));

                var exePath = Path.Combine(Utility.AppDirForRelease(rootAppDirectory, thisRelease), exeName);
                var fileVerInfo = FileVersionInfo.GetVersionInfo(exePath);

                var ret = new Dictionary<ShortcutLocation, ShellLink>();
                foreach (var f in (ShortcutLocation[]) Enum.GetValues(typeof(ShortcutLocation))) {
                    if (!locations.HasFlag(f)) continue;

                    var file = linkTargetForVersionInfo(f, zf, fileVerInfo);

                    this.Log().Info("Creating shortcut for {0} => {1}", exeName, file);

                    ShellLink sl;
                    sl = new ShellLink {
                        Target = updateExe,
                        IconPath = exePath,
                        IconIndex = 0,
                        WorkingDirectory = Path.GetDirectoryName(exePath),
                        Description = zf.Description,
                        Arguments = "--processStart " + exeName,
                    };

                    if (!String.IsNullOrWhiteSpace(programArguments)) {
                        sl.Arguments += String.Format(" -a \"{0}\"", programArguments);
                    }

                    sl.SetAppUserModelId(String.Format("com.squirrel.{0}.{1}", zf.Id, exeName.Replace(".exe", "")));
                    ret.Add(f, sl);
                }

                return ret;
            }

            public void CreateShortcutsForExecutable(string exeName, ShortcutLocation locations, bool updateOnly, string programArguments, string icon)
            {
                this.Log().Info("About to create shortcuts for {0}, rootAppDir {1}", exeName, rootAppDirectory);

                var releases = Utility.LoadLocalReleases(Utility.LocalReleaseFileForAppDir(rootAppDirectory));
                var thisRelease = Utility.FindCurrentVersion(releases);
                var updateExe = Path.Combine(rootAppDirectory, "update.exe");

                var zf = new ZipPackage(Path.Combine(
                    Utility.PackageDirectoryForAppDir(rootAppDirectory),
                    thisRelease.Filename));

                var exePath = Path.Combine(Utility.AppDirForRelease(rootAppDirectory, thisRelease), exeName);
                var fileVerInfo = FileVersionInfo.GetVersionInfo(exePath);

                foreach (var f in (ShortcutLocation[]) Enum.GetValues(typeof(ShortcutLocation))) {
                    // If we are doing update, might as well check all the places.
                    if (!updateOnly && !locations.HasFlag(f)) {
                        continue;
                    }

                    var file = linkTargetForVersionInfo(f, zf, fileVerInfo);
                    var fileExists = File.Exists(file);

                    // NB: If we've already installed the app, but the shortcut
                    // is no longer there, we have to assume that the user didn't
                    // want it there and explicitly deleted it, so we shouldn't
                    // annoy them by recreating it.
                    if (!fileExists && updateOnly) {
                        if (locations.HasFlag(f)) {
                            this.Log().Warn("Wanted to update shortcut {0} but it appears user deleted it", file);
                        }
                        continue;
                    }

                    this.Log().Info("Creating shortcut for {0} => {1}", exeName, file);

                    ShellLink sl;
                    this.ErrorIfThrows(() => {
                        File.Delete(file);

                        sl = new ShellLink {
                            Target = updateExe,
                            IconPath = icon ?? exePath,
                            IconIndex = 0,
                            WorkingDirectory = Path.GetDirectoryName(exePath),
                            Description = zf.Description,
                            Arguments = "--processStart " + exeName,
                        };

                        if (!String.IsNullOrWhiteSpace(programArguments)) {
                            sl.Arguments += String.Format(" -a \"{0}\"", programArguments);
                        }

                        sl.SetAppUserModelId(String.Format("com.squirrel.{0}.{1}", zf.Id, exeName.Replace(".exe", "")));

                        this.Log().Info("About to save shortcut: {0} (target {1}, workingDir {2}, args {3})", file, sl.Target, sl.WorkingDirectory, sl.Arguments);
                        if (ModeDetector.InUnitTestRunner() == false) sl.Save(file);
                    }, "Can't write shortcut: " + file);
                }

                fixPinnedExecutables(zf.Version.Version);
            }

            public void RemoveShortcutsForExecutable(string exeName, ShortcutLocation locations)
            {
                var releases = Utility.LoadLocalReleases(Utility.LocalReleaseFileForAppDir(rootAppDirectory));
                var thisRelease = Utility.FindCurrentVersion(releases);

                var zf = new ZipPackage(Path.Combine(
                    Utility.PackageDirectoryForAppDir(rootAppDirectory),
                    thisRelease.Filename));

                var fileVerInfo = FileVersionInfo.GetVersionInfo(
                    Path.Combine(Utility.AppDirForRelease(rootAppDirectory, thisRelease), exeName));

                foreach (var f in (ShortcutLocation[]) Enum.GetValues(typeof(ShortcutLocation))) {
                    if (!locations.HasFlag(f)) continue;

                    var file = linkTargetForVersionInfo(f, zf, fileVerInfo);

                    this.Log().Info("Removing shortcut for {0} => {1}", exeName, file);

                    this.ErrorIfThrows(() => {
                        if (File.Exists(file)) File.Delete(file);
                    }, "Couldn't delete shortcut: " + file);
                }

                fixPinnedExecutables(zf.Version.Version);
            }

            Task<string> installPackageToAppDir(UpdateInfo updateInfo, ReleaseEntry release)
            {
                return Task.Run(async () => {
                    var zipper = new FastZip();
                    var target = getDirectoryForRelease(release.Version);

                    // NB: This might happen if we got killed partially through applying the release
                    if (target.Exists) {
                        this.Log().Warn("Found partially applied release folder, killing it: " + target.FullName);
                        await Utility.DeleteDirectory(target.FullName);
                    }

                    target.Create();

                    this.Log().Info("Writing files to app directory: {0}", target.FullName);
                    zipper.ExtractZip(
                        Path.Combine(updateInfo.PackageDirectory, release.Filename),
                        target.FullName, FastZip.Overwrite.Always, (o) => true, null, @"lib", true);

                    // Move all of the files out of the lib/ dirs in the NuGet package
                    // into our target App directory.
                    //
                    // NB: We sort this list in order to guarantee that if a Net20
                    // and a Net40 version of a DLL get shipped, we always end up
                    // with the 4.0 version.
                    var libDir = target.GetDirectories().First(x => x.Name.Equals("lib", StringComparison.OrdinalIgnoreCase));
                    var toMove = libDir.GetDirectories().OrderBy(x => x.Name);

                    toMove.ForEach(ld => {
                        ld.GetDirectories()
                            .ForEachAsync(subdir => subdir.MoveTo(subdir.FullName.Replace(ld.FullName, target.FullName)))
                            .Wait();

                        ld.GetFiles()
                            .ForEachAsync(file => {
                                var tgt = Path.Combine(target.FullName, file.Name);
                                this.Log().Info("Moving file {0} to {1}", file.FullName, tgt);
                                if (File.Exists(tgt)) {
                                    Utility.DeleteFileHarder(tgt, true);
                                }
                                file.MoveTo(tgt);
                            })
                            .Wait();
                    });

                    await Utility.DeleteDirectory(libDir.FullName);
                    return target.FullName;
                });
            }

            async Task<ReleaseEntry> createFullPackagesFromDeltas(IEnumerable<ReleaseEntry> releasesToApply, ReleaseEntry currentVersion)
            {
                Contract.Requires(releasesToApply != null);

                // If there are no remote releases at all, bail
                if (!releasesToApply.Any()) {
                    return null;
                }

                // If there are no deltas in our list, we're already done
                if (releasesToApply.All(x => !x.IsDelta)) {
                    return releasesToApply.MaxBy(x => x.Version).FirstOrDefault();
                }

                if (!releasesToApply.All(x => x.IsDelta)) {
                    throw new Exception("Cannot apply combinations of delta and full packages");
                }

                // Smash together our base full package and the nearest delta
                var ret = await Task.Run(() => {
                    var basePkg = new ReleasePackage(Path.Combine(rootAppDirectory, "packages", currentVersion.Filename));
                    var deltaPkg = new ReleasePackage(Path.Combine(rootAppDirectory, "packages", releasesToApply.First().Filename));

                    var deltaBuilder = new DeltaPackageBuilder(Directory.GetParent(this.rootAppDirectory).FullName);

                    return deltaBuilder.ApplyDeltaPackage(basePkg, deltaPkg,
                        Regex.Replace(deltaPkg.InputPackageFile, @"-delta.nupkg$", ".nupkg", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
                });

                if (releasesToApply.Count() == 1) {
                    return ReleaseEntry.GenerateFromFile(ret.InputPackageFile);
                }

                var fi = new FileInfo(ret.InputPackageFile);
                var entry = ReleaseEntry.GenerateFromFile(fi.OpenRead(), fi.Name);

                // Recursively combine the rest of them
                return await createFullPackagesFromDeltas(releasesToApply.Skip(1), entry);
            }

            void executeSelfUpdate(Version currentVersion)
            {
                var targetDir = getDirectoryForRelease(currentVersion);
                var newSquirrel = Path.Combine(targetDir.FullName, "Squirrel.exe");
                if (!File.Exists(newSquirrel)) {
                    return;
                }

                // If we're running in the context of Update.exe, we can't 
                // update ourselves. Instead, ask the new Update.exe to do it
                // once we exit
                var us = Assembly.GetEntryAssembly();
                if (us != null && Path.GetFileName(us.Location).Equals("update.exe", StringComparison.OrdinalIgnoreCase)) {
                    var appName = targetDir.Parent.Name;

                    var p = Process.Start(newSquirrel, "--updateSelf=" + us.Location);
                    this.Log().Info("Started updateSelf pid {0}", p.Id);
                    return;
                }

                // If we're *not* Update.exe, this is easy, it's just a file copy
                Utility.Retry(() =>
                    File.Copy(newSquirrel, Path.Combine(targetDir.Parent.FullName, "Update.exe"), true));
            }

            async Task invokePostInstall(Version currentVersion, bool isInitialInstall, bool firstRunOnly, bool silentInstall)
            {
                var targetDir = getDirectoryForRelease(currentVersion);
                var args = isInitialInstall ?
                    String.Format("--squirrel-install {0}", currentVersion) :
                    String.Format("--squirrel-updated {0}", currentVersion);

                var squirrelApps = SquirrelAwareExecutableDetector.GetAllSquirrelAwareApps(targetDir.FullName);

                this.Log().Info("Squirrel Enabled Apps: [{0}]", String.Join(",", squirrelApps));

                // For each app, run the install command in-order and wait
                if (!firstRunOnly) await squirrelApps.ForEachAsync(async exe => {
                    using (var cts = new CancellationTokenSource()) { 
                        cts.CancelAfter(15 * 1000);
                        try {
                            await Utility.InvokeProcessAsync(exe, args, cts.Token);
                        } catch (Exception ex) {
                            this.Log().ErrorException("Couldn't run Squirrel hook, continuing: " + exe, ex);
                        }
                    }
                }, 1 /* at a time */);

                // If this is the first run, we run the apps with first-run and 
                // *don't* wait for them, since they're probably the main EXE
                if (squirrelApps.Count == 0) {
                    this.Log().Warn("No apps are marked as Squirrel-aware! Going to run them all");

                    squirrelApps = targetDir.EnumerateFiles()
                        .Where(x => x.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        .Where(x => !x.Name.StartsWith("squirrel.", StringComparison.OrdinalIgnoreCase))
                        .Select(x => x.FullName)
                        .ToList();

                    // Create shortcuts for apps automatically if they didn't
                    // create any Squirrel-aware apps
                    squirrelApps.ForEach(x => CreateShortcutsForExecutable(Path.GetFileName(x), ShortcutLocations.Defaults, isInitialInstall == false, null, null));
                }

                if (!isInitialInstall || silentInstall) return;

                var firstRunParam = isInitialInstall ? "--squirrel-firstrun" : "";
                squirrelApps
                    .Select(exe => new ProcessStartInfo(exe, firstRunParam) { WorkingDirectory = Path.GetDirectoryName(exe) })
                    .ForEach(info => {
                        var p = Process.Start(info);
                        this.Log().Info("ran {0}, pid {1}", info.FileName, p.Id);
                    });
            }

            void fixPinnedExecutables(Version newCurrentVersion) 
            {
                if (Environment.OSVersion.Version < new Version(6, 1)) {
                    this.Log().Warn("fixPinnedExecutables: Found OS Version '{0}', exiting...", Environment.OSVersion.VersionString);
                    return;
                }

                var newCurrentFolder = "app-" + newCurrentVersion;

                this.Log().Info("fixPinnedExecutables: newCurrentFolder: {0}", newCurrentFolder);

                var newAppPath = Path.Combine(rootAppDirectory, newCurrentFolder);
                bool newVersionExists = Directory.Exists(newAppPath);

                Func<FileInfo, ShellLink> resolveLink = file => {
                    try {
                        return new ShellLink(file.FullName);
                    } catch (Exception ex) {
                        var message = String.Format("File '{0}' could not be converted into a valid ShellLink", file.FullName);
                        this.Log().WarnException(message, ex);
                        return null;
                    }
                };

                var taskbarPath = new Tuple<string, SearchOption>(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar"),
                    SearchOption.TopDirectoryOnly);
                var startMenuPath = new Tuple<string, SearchOption>(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
                    SearchOption.AllDirectories);
                var desktopPath = new Tuple<string, SearchOption>(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                    SearchOption.TopDirectoryOnly);
                Tuple<string, SearchOption>[] pathsToUpdate = { taskbarPath, startMenuPath, desktopPath };

                foreach (var searchPath in pathsToUpdate) {
                    var shellLinks = (new DirectoryInfo(searchPath.Item1)).GetFiles("*.lnk", searchPath.Item2)
                        .Select(resolveLink)
                        .Where(x => x != null)
                        .ToArray();

                    foreach (var shortcut in shellLinks) {
                        try {
                            updateLink(shortcut, newAppPath, newVersionExists);
                        }
                        catch (Exception ex) {
                            var message = String.Format("fixPinnedExecutables: shortcut failed: {0}", shortcut.Target);
                            this.Log().ErrorException(message, ex);
                        }
                    }
                }
            }

            string updatePath(string pathToUpdate, string newAppPath)
            {
                if (!pathToUpdate.StartsWith(rootAppDirectory)) {
                    return pathToUpdate;
                }
                string afterRoot = pathToUpdate.Substring(rootAppDirectory.Length + 1);
                string[] pathParts = afterRoot.Split(Path.DirectorySeparatorChar);
                // if target is in app subdir, update it
                if (pathParts[0].StartsWith("app-")) {
                    pathParts[0] = newAppPath;
                    pathToUpdate = Path.Combine(pathParts);
                }
                return pathToUpdate;
            }

            void updateLink(ShellLink shortcut, string newAppPath, bool newVersionExists)
            {
                this.Log().Info("Processing shortcut '{0}'", shortcut.Target);

                string expectedStart = rootAppDirectory + "\\app-";

                if (!shortcut.WorkingDirectory.StartsWith(expectedStart, StringComparison.OrdinalIgnoreCase)) {
                    this.Log().Info("'{0}' is not in '{1}', skipping", shortcut.WorkingDirectory, rootAppDirectory);
                    return;
                }

                if (!newVersionExists) {
                    this.Log().Info("Unpinning {0} from taskbar", shortcut.Target);
                    TaskbarHelper.UnpinFromTaskbar(shortcut.Target);
                    return;
                }

                shortcut.Target = updatePath(shortcut.Target, newAppPath);
                shortcut.WorkingDirectory = updatePath(shortcut.WorkingDirectory, newAppPath);
                shortcut.IconPath = updatePath(shortcut.IconPath, newAppPath);

                this.Log().Info("Updating shortcut to {0}", shortcut.Target);
                shortcut.Save();
            }

            // NB: Once we uninstall the old version of the app, we try to schedule
            // it to be deleted at next reboot. Unfortunately, depending on whether
            // the user has admin permissions, this can fail. So as a failsafe,
            // before we try to apply any update, we assume previous versions in the
            // directory are "dead" (i.e. already uninstalled, but not deleted), and
            // we blow them away. This is to make sure that we don't attempt to run
            // an uninstaller on an already-uninstalled version.
            async Task cleanDeadVersions(Version originalVersion, Version currentVersion, bool forceUninstall = false)
            {
                if (currentVersion == null) return;

                var di = new DirectoryInfo(rootAppDirectory);
                if (!di.Exists) return;

                this.Log().Info("cleanDeadVersions: for version {0}", currentVersion);

                string originalVersionFolder = null;
                if (originalVersion != null) {
                    originalVersionFolder = getDirectoryForRelease(originalVersion).Name;
                    this.Log().Info("cleanDeadVersions: exclude original version folder {0}", originalVersionFolder);
                }

                string currentVersionFolder = null;
                if (currentVersion != null) {
                    currentVersionFolder = getDirectoryForRelease(currentVersion).Name;
                    this.Log().Info("cleanDeadVersions: exclude current version folder {0}", currentVersionFolder);
                }

                // NB: If we try to access a directory that has already been 
                // scheduled for deletion by MoveFileEx it throws what seems like
                // NT's only error code, ERROR_ACCESS_DENIED. Squelch errors that
                // come from here.
                var toCleanup = di.GetDirectories()
                    .Where(x => x.Name.ToLowerInvariant().Contains("app-"))
                    .Where(x => x.Name != currentVersionFolder && x.Name != originalVersionFolder)
                    .Where(x => !isAppFolderDead(x.FullName));

                if (forceUninstall == false) {
                    await toCleanup.ForEachAsync(async x => {
                        var squirrelApps = SquirrelAwareExecutableDetector.GetAllSquirrelAwareApps(x.FullName);
                        var args = String.Format("--squirrel-obsolete {0}", x.Name.Replace("app-", ""));

                        if (squirrelApps.Count > 0) {
                            // For each app, run the install command in-order and wait
                            await squirrelApps.ForEachAsync(async exe => {
                                using (var cts = new CancellationTokenSource()) { 
                                    cts.CancelAfter(10 * 1000);

                                    try {
                                        await Utility.InvokeProcessAsync(exe, args, cts.Token);
                                    } catch (Exception ex) {
                                        this.Log().ErrorException("Couldn't run Squirrel hook, continuing: " + exe, ex);
                                    }
                                }
                            }, 1 /* at a time */);
                        }
                    });
                }

                // Include dead folders in folders to :fire:
                toCleanup = di.GetDirectories()
                    .Where(x => x.Name.ToLowerInvariant().Contains("app-"))
                    .Where(x => x.Name != currentVersionFolder && x.Name != originalVersionFolder);

                // Finally, clean up the app-X.Y.Z directories
                await toCleanup.ForEachAsync(async x => {
                    try {
                        await Utility.DeleteDirectoryWithFallbackToNextReboot(x.FullName);

                        if (Directory.Exists(x.FullName)) {
                            // NB: If we cannot clean up a directory, we need to make 
                            // sure that anyone finding it later won't attempt to run
                            // Squirrel events on it. We'll mark it with a .dead file
                            markAppFolderAsDead(x.FullName);
                        }
                    } catch (UnauthorizedAccessException ex) {
                        this.Log().WarnException("Couldn't delete directory: " + x.FullName, ex);

                        // NB: Same deal as above
                        markAppFolderAsDead(x.FullName);
                    }
                });

                // Clean up the packages directory too
                var releasesFile = Utility.LocalReleaseFileForAppDir(rootAppDirectory);
                var entries = ReleaseEntry.ParseReleaseFile(File.ReadAllText(releasesFile, Encoding.UTF8));
                var pkgDir = Utility.PackageDirectoryForAppDir(rootAppDirectory);
                var releaseEntry = default(ReleaseEntry);

                foreach (var entry in entries) {
                    if (entry.Version == currentVersion) {
                        releaseEntry = ReleaseEntry.GenerateFromFile(Path.Combine(pkgDir, entry.Filename));
                        continue;
                    }

                    File.Delete(Path.Combine(pkgDir, entry.Filename));
                }

                ReleaseEntry.WriteReleaseFile(new[] { releaseEntry }, releasesFile);
            }

            static void markAppFolderAsDead(string appFolderPath)
            {
                File.WriteAllText(Path.Combine(appFolderPath, ".dead"), "");
            }

            static bool isAppFolderDead(string appFolderPath)
            {
                return File.Exists(Path.Combine(appFolderPath, ".dead"));
            }

            internal async Task<List<ReleaseEntry>> updateLocalReleasesFile()
            {
                return await Task.Run(() => ReleaseEntry.BuildReleasesFile(Utility.PackageDirectoryForAppDir(rootAppDirectory)));
            }

            IEnumerable<DirectoryInfo> getReleases()
            {
                var rootDirectory = new DirectoryInfo(rootAppDirectory);

                if (!rootDirectory.Exists) return Enumerable.Empty<DirectoryInfo>();

                return rootDirectory.GetDirectories()
                    .Where(x => x.Name.StartsWith("app-", StringComparison.InvariantCultureIgnoreCase));
            }

            DirectoryInfo getDirectoryForRelease(Version releaseVersion)
            {
                return new DirectoryInfo(Path.Combine(rootAppDirectory, "app-" + releaseVersion));
            }

            string linkTargetForVersionInfo(ShortcutLocation location, IPackage package, FileVersionInfo versionInfo)
            {
                var possibleProductNames = new[] {
                    versionInfo.ProductName,
                    package.Title,
                    versionInfo.FileDescription,
                    Path.GetFileNameWithoutExtension(versionInfo.FileName)
                };

                var possibleCompanyNames = new[] {
                    versionInfo.CompanyName,
                    package.Authors.FirstOrDefault() ?? package.Id,
                };

                var prodName = possibleCompanyNames.First(x => !String.IsNullOrWhiteSpace(x));
                var pkgName = possibleProductNames.First(x => !String.IsNullOrWhiteSpace(x));

                return getLinkTarget(location, pkgName, prodName);
            }

            string getLinkTarget(ShortcutLocation location, string title, string applicationName, bool createDirectoryIfNecessary = true)
            {
                var dir = default(string);

                switch (location) {
                case ShortcutLocation.Desktop:
                    dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    break;
                case ShortcutLocation.StartMenu:
                    dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs", applicationName);
                    break;
                case ShortcutLocation.Startup:
                    dir = Environment.GetFolderPath (Environment.SpecialFolder.Startup);
                    break;
                case ShortcutLocation.AppRoot:
                    dir = rootAppDirectory;
                    break;
                case ShortcutLocation.Taskbar:
                    dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Internet Explorer", "Quick Launch", "User Pinned", "TaskBar");
                    break;
                }

                if (createDirectoryIfNecessary && !Directory.Exists(dir)) {
                    Directory.CreateDirectory(dir);
                }

                return Path.Combine(dir, title + ".lnk");
            }
        }
    }
}
