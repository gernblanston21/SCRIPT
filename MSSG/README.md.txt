# MSSG Node.js Odds Service

## Overview

[cite_start]This package contains a Node.js application (`server.js`) designed to run as a background Windows service[cite: 1]. [cite_start]Its primary purpose is to continuously fetch live sports odds and prop bet data from the FanDuel API[cite: 1]. The data is saved as JSON files into a `cache` directory, where they can be read by a Viz Trio VTW (Template Wizard) script to populate graphics.

The service is managed by a series of `.bat` script files for easy installation, removal, and control.

## Installation and Uninstallation

**Prerequisites:**
* [cite_start]**Node.js:** The Node.js runtime must be installed on the system and accessible via the system's PATH, or placed in a `bin/node` folder within this directory[cite: 3].
* **NSSM (Non-Sucking Service Manager):** `nssm.exe` is required to manage the Windows service. [cite_start]It should be placed in a `bin/nssm` folder within this directory[cite: 2].

**Configuration:**
Before installing, you must create or edit the `.env` file in the main directory. [cite_start]It needs to contain your FanDuel API key like this[cite: 9]:
`FD_API_KEY=YOUR_API_KEY_HERE`

### To Install the Service:

1.  Make sure the prerequisites above are met.
2.  Right-click on the `install_service.bat` file.
3.  Select **"Run as administrator"**.
4.  [cite_start]The script will install the service named `MSSG-Node` and start it automatically[cite: 1, 6].

### To Uninstall the Service:

1.  Right-click on the `uninstall_service.bat` file.
2.  Select **"Run as administrator"**.
3.  [cite_start]The script will stop and completely remove the `MSSG-Node` service from the system[cite: 13].

## Script (Batch File) Descriptions

These files are used to manage and test the service.

* `install_service.bat`
    * [cite_start]**Purpose:** Installs the `server.js` application as a Windows service that starts automatically with the computer[cite: 1]. [cite_start]It also configures logging and log rotation[cite: 5].
    * **Requires:** Administrator privileges.

* `uninstall_service.bat`
    * [cite_start]**Purpose:** Stops and permanently removes the Windows service from the system[cite: 13].
    * **Requires:** Administrator privileges.

* `start_service.bat`
    * [cite_start]**Purpose:** Manually starts the service if it has been stopped[cite: 8].

* `stop_service.bat`
    * [cite_start]**Purpose:** Manually stops the service without removing it[cite: 10]. The service will not run again until it is started with `start_service.bat` or the computer is rebooted.

* `test_run_console.bat`
    * **Purpose:** This is a **debugging tool**. [cite_start]It runs the `server.js` application directly in a console window instead of as a hidden background service[cite: 11, 12]. This is extremely useful for troubleshooting, as you can see all the live log output and any potential errors in real-time. **Do not** run this if the service is already running.

## Pausing the Data Feed

There are two ways to stop the data feed: a "soft stop" (Pause) and a "hard stop" (Stop Service).

### Method 1: The "Pause Cache" Button (Recommended)

The VTW control panel has a "Pause Cache" button. This is the safest way to temporarily halt the data feed during a broadcast.

* **How it Works:** When you click "Pause Cache," the VTW script creates an empty file named `STOP.flag` in the `cache` folder.
* **What Happens:** The `server.js` service is programmed to look for this file every second. When it sees `STOP.flag`, it enters a PAUSED state.
    * It immediately stops making any new calls to the FanDuel API.
    * To prevent stale data from being used, it rewrites the odds and lines in all `liveCore.json` files to "N/A", while preserving the team names. It also blanks the prop bet files.
* **Resuming:** When you click "Resume Cache," the VTW script deletes the `STOP.flag` file. The service detects this, resumes its normal polling schedule, and repopulates the JSON files with live data.

### Method 2: Using `stop_service.bat`

[cite_start]Running `stop_service.bat` performs a "hard stop" by terminating the service entirely[cite: 10].

* **Difference:** Unlike the pause flag, this completely shuts down the `server.js` application. It will not be running in the background and cannot be resumed by deleting the `STOP.flag` file.
* **Use Case:** This should be used for system maintenance, updating the `server.js` script, or shutting down for an extended period. To get data flowing again, you must run `start_service.bat`.

---
---

## Files and Folders to Delete for Delivery

Before packaging your project for delivery to a client or another user, you should clean out any files and folders that are generated during runtime. This ensures they start with a clean slate.

**Delete the following:**

* **`cache/` (the entire folder)**
    * **Reason:** This folder is created and populated by `server.js` at runtime. It contains all the downloaded JSON data and logs. [cite_start]The service will automatically recreate this folder when it starts[cite: 1]. Deleting it ensures you are not delivering old or stale data.

* **`logs/` (if it exists outside the cache folder)**
    * **Reason:** Similar to the cache, this contains runtime logs that are not needed for a clean installation. [cite_start]The `install_service.bat` script configures logs to be written to `cache/logs/`, so deleting the `cache` folder should cover this[cite: 1].

* **`.env`**
    * [cite_start]**Reason:** This file contains your secret API key[cite: 9]. You should **NEVER** deliver your project with your key inside it.
    * **Recommendation:** Either delete the file entirely and instruct the user to create it, or replace your key with a placeholder like `FD_API_KEY=REPLACE_WITH_YOUR_KEY`.