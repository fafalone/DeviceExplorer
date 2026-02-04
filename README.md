# DevExplorer
## Device Explorer v1.3

**A basic implementation of a device manager like the system's in twinBASIC.**

![image](https://github.com/fafalone/DeviceExplorer/assets/7834493/de43691b-7f18-424b-aaf7-e01a4aac3092)

---

**Update (04 Feb 2026):**
'v1.3
- Code fixes for new twinBASIC syntax rules; update WDL package and change to linked.
- Fix Computer picture black background bug
- Existing binary build had broken right click
- Add refreshes for device removals


**Update (04 Sep 2024):** 
'v1.2
- Weird issue with icons being off by one... I commented out lines that adjusted these; but in previous versions they were required to get the correct icon. Still is, for Other Devices.
  Please let me know if you see the wrong icon on any category or device!
 
- Fix for icons disappearing when an item was selected.
'
- Fix for problem running in IDE multiple times without compiler restart.

**Update (21 Jun 2024):** v1.1 The source has been updated to compile with the latest tB version, since breaking changes were introduced in recent builds. Source requires tB 563 or newer to build now. No feature updates. No binary updates.

This project started out as just a proof of concept to list devices and test out some disable/enable code I wanted to try, but I got interested enough to turn it into a full blown application with all of the basic functionality of the system Device Manager. This project was started quite a while back, but I got stuck on a couple things, and with so many other interesting projects in the pipeline, shelved it for a while. But when I came across a solution to one of the problems (showing the device properties popup) recently, I got the motivation to finish it up. You can enable/disable devices, remove them, completely uninstall them, update their drivers, or eject them. To be honest, I'm not clear on the difference or true meaning behind some of these things, as neither Device Manager itself nor the APIs provide particularly good documentation of the details of each option, at least on the surface, I'll dig a little more in the future. 

There's an unusual, for now, reason to do this project in twinBASIC: The enable/disable/remove APIs (at a minimum) do not allow 32bit applications to call them on 64bit Windows-- the `SetupDiCallClassInstaller` API will fail with `ERROR_IN_WOW64`, as [documented on MSDN](https://learn.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdicallclassinstaller). Additionally, there's a 'Resources' tab in the properties:

![image](https://github.com/fafalone/DeviceExplorer/assets/7834493/4bbe012b-f5b5-406a-804f-2b2b90eb78cd)

THis tab will not load under WOW64. I haven't thoroughly tested the rest of the functionality on WOW64, but the bottom line is you'd need extensive workarounds and hacks to accomplish the same tasks in VB6, because a large portion of hardware setup functionality has been disallowed under WOW64. This is the first time I've encountered APIs explicitly blocked like this, but I strongly suspect it won't be the last. (One other slightly similar situation unrelated to this app; a small number of APIs use `_fastcall` on x86; tB doesn't support that either, but the x64 version uses the standard x64 calling convention). Memory access isn't the only reason to move to x64; this is a preview of things to come with Microsoft marking 32bit apps insecure and limiting their access.

### Features

- Remove, disable, enable, uninstall, and eject devices (not all devices support all actions).

- Start the Update Driver wizard for devices.

- Show Device Properties on double click.

- Can list devices that have been installed but aren't present ('hidden devices') like the system device manager.

- Shows devices with a problem with the same overlay icon as the system device manager (including the alternate overlay for when the problem is 'the user disabled it'); loads the problem text from the system.

- Uses techniques from my other projects to run from the IDE while using the resources of the compiled exe.

- Shows devices without a defined class under 'Other devices' like the system does (getting these to show was the other thing I was stuck on for a while).

  ![image](https://github.com/fafalone/DeviceExplorer/assets/7834493/284c5eef-183c-4e5a-b43f-86847a00b2ec)

- Has a shortcut to open the Devices and Printers folder.

- Scan for hardware changes menu option like system device manager.

### Requirements

  - As noted earlier, many features requiring using the 64bit build/IDE mode on 64bit Windows.
 
  - Most features require running as administrator (you can still list the devices while not elevated, but not perform most actions). The app enables the SeLoadDriverPrivilege when elevated as that's required for eject and possibly other features.
 
  - Not likely to work on Windows XP or earlier.
 
### How it works

  First of all, this project again makes heavy use of WinDevLib (formerly tbShellLib), my project to make programming in tB more like having windows.h available in other languages. It has an extensive set of setup APIs from setupapi.dll, newdev.dll, cfgmgr32.dll, and devmgr.dll.


  #### Enumerating devices

  First step is getting a list of all the classes registered on the system:

  ```vba
     ret = SetupDiBuildClassInfoList(0, vbNullPtr, 0, cbReq)
    
    If cbReq > 0 Then
        ReDim arClasses(cbReq - 1)

        ret = SetupDiBuildClassInfoList(0, arClasses(0), UBound(arClasses) + 1, cbReq)
        If ret Then
           For i = 0 To UBound(arClasses) - 1
                   sBufN = String$(MAX_CLASS_NAME_LEN, 0)
                ret = SetupDiClassNameFromGuid(arClasses(i), sBufN, Len(sBufN), cchReq)
                If cchReq Then sBufN = Left$(sBufN, cchReq - 1)
                cchReq = 0
                ret = SetupDiGetClassDescription(arClasses(i), vbNullString, 0, cchReq)
                If cchReq > 0 Then
                    sBufD = String$(cchReq, 0)
                    ret = SetupDiGetClassDescription(arClasses(i), sBufD, Len(sBufD), cchReq)
                End If
  ```

  We then go through the classes and enumerate the devices for each one:

  ```vba
      For i = 0 To UBound(DMSet)
        If DMSet(i).bDevice Then Continue For
        
        'Is a class
        hSet = SetupDiGetClassDevs(DMSet(i).AsscGUID, vbNullString, Me.hWnd, IIf(Check3.Value = vbChecked, 0&, DIGCF_PRESENT))
        If hSet = INVALID_HANDLE_VALUE Then Continue For
            
        tDevInfo.cbSize = LenB(tDevInfo)
        j = 0
        
        Do While SetupDiEnumDeviceInfo(hSet, j, tDevInfo)
            cchReq = 0: cbReq = 0
            sBufN = "": sBufID = "": sBufP = ""
            dwCap = 0: bProblem = False
            dwStatus = 0: nProbCode = 0
            fPresent = 0
            dwMask = 0: dwState = 0
            ret = SetupDiGetDeviceInstanceId(hSet, tDevInfo, vbNullString, 0, cchReq)
            If cchReq Then
                sBufID = String$(cchReq, 0)
                ret = SetupDiGetDeviceInstanceId(hSet, tDevInfo, sBufID, Len(sBufID), cchReq)
                If InStr(sBufID, Chr$(0)) > 1 Then
                    sBufID = Left$(sBufID, InStr(sBufID, Chr$(0)) - 1)
                End If
            End If
  ```

  From there there's a bunch of APIs to query information about each device, like it's friendly name, status, icon, etc, in order to display it. There's way too much to go over in detail, but I think the code is fairly readable, so check out the full project.

  #### Showing property pages

  This one we resort to a little undocumented funcitonality, because the documented way has a bizarre side effect: it will activate DPI awareness for your app if it's off, resulting in the window becoming tiny. So we call the API the documented way is just a thin wrapper for directly:

  ```vba
            'DeviceProperties_RunDLL Me.hWnd, 0, "/DeviceId " & sBufID, SW_SHOW
            'That's the official way, but for some bizarre reason it activates DPI awareness
            'if it's not already on; that makes the window tiny. Calling the API it wraps
            'directly avoids the problem. No, 0 for hwnd/hinst doesn't change anything.
            'Plus here we can add the resources tab; I don't know how to specify that flag
            'for the command line version.
            DevicePropertiesEx Me.hWnd, vbNullString, sBufID, DEVPROP_SHOW_RESOURCE_TAB, 0 'DEVPROP_SHOW_RESOURCE_TAB doesn't seem to work under WOW64
  ```

  #### Enabling/disabling devices

  This is the most complex part. After the standard code to identify the device instance data, we use a `SP_PROPCHANGE_PARAMS` type with `SetupDiSetClassInstallParams` followed by `SetupDiCallClassInstaller`. No idea why MS made that two different functions.

  ```vba
  Dim tParams As SP_PROPCHANGE_PARAMS

             tParams.ClassInstallHeader.cbSize = LenB(Of SP_CLASSINSTALL_HEADER)
            tParams.ClassInstallHeader.InstallFunction = DIF_PROPERTYCHANGE
            If fEnable Then
                tParams.StateChange = DICS_ENABLE
            Else
                tParams.StateChange = DICS_DISABLE
            End If
            tParams.Scope = DICS_FLAG_CONFIGSPECIFIC
            ret = SetupDiSetClassInstallParams(hSet, tDevInfo, tParams, LenB(Of SP_PROPCHANGE_PARAMS))
            If ret Then
                ret = SetupDiCallClassInstaller(DIF_PROPERTYCHANGE, hSet, tDevInfo)
                If ret Then
                    SetItemOverlayIndex(DMSet(idx).hItem, IIf(fEnable, 0, 3))
                    SetupDiDestroyDeviceInfoList hSet
                    Return S_OK
```

Remove is very similar, just with a different function and params type.


The are fairly straightforward so I'll conclude here; definitely let me know if there's any questions, comments, bugs, or feature requests! 

PS- No making fun of my graphics card please :) I got a 6750XT for Christmas, but haven't installed it yet.

