# Desktop Connector "Unable to register drive" — Root Cause Analysis

**BriComp IT Consulting Services** — [bricomp.com](https://bricomp.com)

This is the diagnostic path behind [Repair-BCDesktopConnectorSyncRoot.ps1](Repair-BCDesktopConnectorSyncRoot.ps1). It is written out in full because the failure is easy to misdiagnose — every obvious remedy fails, and the one piece of evidence that explains it sits in a registry hive most people never look at.

All examples below use placeholder values.

---

## The presentation

Autodesk Desktop Connector launches, appears in the system tray, sits there for roughly ten seconds, and vanishes. No error dialog. Re-launching produces the same result. The machine in the originating case had been migrated from Entra (Azure AD) join to on-premises Active Directory using a third-party migration tool.

Things that did **not** fix it:

- Reinstalling Desktop Connector
- Creating a brand new Windows profile for the user
- Signing out and back in

That second one matters, and we will come back to why it could never have worked.

---

## Reading the log

Desktop Connector writes newline-delimited JSON to:

```
%LOCALAPPDATA%\Autodesk\Desktop Connector\Logs
```

Filtering for anything above `Information` across several startups produced the same three-line story every time:

```
Mounting virtual file system for "<drive name>"
    Workspace: C:\Users\<user>\DC\ACCDocs

"<drive name>" is not registered.

Failed to register virtual file system for "<drive name>"
CloudFileApi.CloudFileApiException: Access is denied.
    (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))
 ---> System.UnauthorizedAccessException
   at Windows.Storage.Provider.StorageProviderSyncRootManager.Register(StorageProviderSyncRootInfo)
   at Comet.VirtualFileSystem.CloudFileApi.Registrar.RegisterSyncRoot(...)

Tray Shutdown - ExitSource = "Unable to register drive"
```

Two observations shape everything that follows.

**First, this is not an Autodesk problem.** The same log shows `UserLoginState: LoggedIn`, successful token refreshes, and successful `GetHubsAsync` calls against Autodesk's cloud. Authentication, licensing and network are all healthy. The process dies on a purely local Windows call.

**Second, `StorageProviderSyncRootManager.Register` is a Windows API, not an Autodesk one.** It is the Cloud Files API — the same mechanism OneDrive and Dropbox use for on-demand files. An `E_ACCESSDENIED` out of it is a Windows registration failure that Desktop Connector is merely reporting.

---

## Where sync roots actually live

Per Microsoft's cloud storage provider documentation, registration writes to:

```
HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager
```

and the key name follows the pattern:

```
<provider ID>!<Windows SID>!<account ID>
```

**The user's SID is part of the key name**, and a `UserSyncRoots\<SID>` value underneath records the folder that registration claims on disk.

This single fact explains why creating a new Windows profile was useless. These registrations are machine-wide, in `HKLM`, keyed by SID. A new profile gives the user a new profile folder and a fresh `HKCU` — but the same SID, and no visibility into the hive that actually matters.

---

## Enumerating the hive

```powershell
$b = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager'
Get-ChildItem $b | Select-Object -ExpandProperty PSChildName
```

On the affected machine this returned four entries — two providers, each appearing twice:

```
DesktopConnector!QXV0b2Rlc2sgRG9jcw==!S-1-12-1-<...>!User
DesktopConnector!QXV0b2Rlc2sgRG9jcw==!S-1-5-21-<...>!User
OneDrive!S-1-12-1-<...>!Business1|<tenant guid>
OneDrive!S-1-5-21-<...>!Business1|<tenant guid>
```

Two details are worth decoding.

**The middle segment is base64.** `QXV0b2Rlc2sgRG9jcw==` decodes to `Autodesk Docs`.

**`S-1-12-1-…` is an Entra ID SID.** Entra does not issue domain-SID-plus-RID the way on-premises AD does. It takes the object's GUID, splits the sixteen bytes into four 32-bit unsigned integers, and uses those as the trailing sub-authorities. That SID *is* the Entra object ID, re-encoded. `S-1-5-21-…` is the on-premises AD account the user now signs in with.

So the hive contained both identity eras side by side — the pre-migration Entra identity and the post-migration AD identity — for every cloud provider on the box.

---

## Ruling out the obvious

Before assuming stale data, check whether the permissions are simply broken:

```powershell
Get-Acl $b | Format-List
```

On the affected machine this came back completely standard — `BUILTIN\Users` holding `SetValue, CreateSubKey, ReadKey`, which is the stock ACL. Nothing had been mangled by the migration. A migration that re-ACLs the registry is a reasonable hypothesis, and it was wrong here.

The presence of a working `OneDrive!S-1-5-21-…` entry is the confirming evidence: the live account *can* register sync roots on this machine. Whatever is failing is specific to this one registration.

---

## The actual cause

Inspect what each Desktop Connector key claims:

```powershell
Get-ChildItem $b | Where-Object PSChildName -like 'DesktopConnector!*' | ForEach-Object {
    $p = "$b\$($_.PSChildName)"
    [pscustomobject]@{
        Key   = $_.PSChildName
        Owner = (Get-Acl $p).Owner
        Paths = (Get-ItemProperty "$p\UserSyncRoots" -ErrorAction SilentlyContinue |
                 Select-Object * -Exclude PS* | Out-String).Trim()
    }
} | Format-List
```

The result was asymmetric, and that asymmetry is the whole answer:

| Key | `UserSyncRoots` |
|---|---|
| `…!S-1-12-1-…!User` (dead Entra identity) | `C:\Users\<user>\DC\ACCDocs` |
| `…!S-1-5-21-…!User` (live AD account) | *empty* |

The registration belonging to an identity that no longer exists on the machine still held the claim on the workspace folder. The live account's registration got far enough to create its key, then failed before it could claim anything.

Microsoft's `CfRegisterSyncRoot` documentation states the rule plainly:

> No two sync root trees are allowed to overlap. […] The platform is responsible for persistently remembering all sync roots registered on a given volume, and failing any attempt to create overlapping sync roots.

Windows was doing exactly what it is documented to do. The migration changed the user's SID; the old sync root registration was not cleaned up; it kept its claim on `ACCDocs`; and every subsequent attempt by the live account to register that same folder was refused.

Confirming the identity really is dead — the old SID should have no profile entry:

```powershell
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' |
    Select-Object -ExpandProperty PSChildName
```

On the affected machine the Entra SID was absent, confirming the migration had re-pointed the profile to the new SID and left the sync root registration behind. This is the test the repair script automates.

---

## The fix

1. Export `SyncRootManager` as a backup.
2. Exit Desktop Connector.
3. Delete the orphaned `DesktopConnector!…!<dead SID>!User` key, and the empty one for the live SID.
4. Delete the leftover Explorer shell namespace entries for each key's `NamespaceCLSID` value:
   - `HKU\<sid>\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{CLSID}`
   - `HKU\<sid>\Software\Classes\CLSID\{CLSID}`
5. Restart.

One subtlety worth knowing: clean the namespace entries out of **every loaded user hive**, not just the hive matching the key's SID. Explorer creates that node under whichever user is rendering the sidebar, which is not necessarily the user the sync root is keyed to. In the originating case the dead key was keyed to the Entra SID while the ghost node lived in the live account's `HKCU`.

Leave healthy registrations for other providers alone.

---

## Why not just uninstall and reinstall?

Because uninstalling Desktop Connector does not remove a sync root belonging to a SID that no longer resolves. The reinstall walks into the same collision. This is why the problem survives every remedy that operates at the application layer.

---

## A note on the diagnostic order

The useful sequence here was:

1. **Read the exception, not the symptom.** "Desktop Connector won't start" is unactionable; `StorageProviderSyncRootManager.Register` returning `E_ACCESSDENIED` points at exactly one subsystem.
2. **Check what still works.** Successful sign-in and cloud calls in the same log eliminated authentication, licensing and network in one step.
3. **Test the permission hypothesis before assuming it.** The parent key ACL was healthy, which killed the most intuitive theory early rather than after hours of re-ACLing.
4. **Look for a working peer.** OneDrive registering successfully under the same SID proved the failure was specific, not systemic.
5. **Compare claims, not just keys.** The key names alone looked like harmless duplication. Only the `UserSyncRoots` values revealed which key held the contested path.

---

## References

- [Integrate a Cloud Storage Provider](https://learn.microsoft.com/en-us/windows/win32/shell/integrate-cloud-storage) — sync root registry location and key naming
- [CfRegisterSyncRoot](https://learn.microsoft.com/en-us/windows/win32/api/cfapi/nf-cfapi-cfregistersyncroot) — the non-overlapping sync root rule and access requirements
- [StorageProviderSyncRootManager.Register](https://learn.microsoft.com/en-us/uwp/api/windows.storage.provider.storageprovidersyncrootmanager.register) — the API Desktop Connector calls
- [About Autodesk Forma Connector](https://help.autodesk.com/cloudhelp/ENU/CONNECT-User-Guide/files/About_Autodesk_Docs_Connector.htm) — the Autodesk Docs to Forma rebrand, and confirmation that local paths such as `ACCDocs` were left unchanged
