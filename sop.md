# AVD Environment Health Reporting Script SOP

## Overview

The AVD-EnvironmentHealth tool is a single, read-only PowerShell script designed to assess an Azure Virtual Desktop (AVD) environment end-to-end. This includes evaluating Host Pools, Session Hosts, Health Checks, and Profile Storage. The script is assessment-only and makes no write or modify calls against any Azure resource.

## Prerequisites

* **PowerShell:** Version 5.1 or 7.x.


* **Modules:** `Az.Accounts`, `Az.DesktopVirtualization`, `Az.Compute`, `Az.Monitor`, `Az.Resources`, `Az.Network`, `Az.Storage`, `Az.Security`, and `Az.NetAppFiles`. All of these modules are required regardless of the target storage backend.


* **Permissions:** Reader role (minimum) on the AVD resource group and the profile storage resource group. No write permissions are required. Do not assign the Contributor role to run this tool.


* **Authentication:** An authenticated Az context via interactive sign-in, device code, service principal, or managed identity.


* **Network:** Outbound HTTPS (443) access from the execution host to Azure Resource Manager and Marketplace endpoints.



## Configuration

The FSLogix profile storage backend varies by environment and must always be configured prior to running the assessment.

1. Confirm with the environment owner whether FSLogix profiles are stored on Azure NetApp Files (ANF), Azure Files, or both.


2. Copy the configuration template: `config.template.json` → `config.<client>-<env>.json`.


3. Populate the client name, environment name, subscription/tenant ID, and AVD resource group.


4. Set the `profileStorageType` variable to `ANF`, `AzureFiles`, or `Both`.


5. Populate the corresponding account and resource group fields (e.g., `anfResourceGroupName`, `storageResourceGroupName`, `storageAccountName`) based on your backend choice.


6. **Security Warning:** Store any configuration files containing subscription or tenant IDs outside of source control.



## Execution Steps

### 1. Authenticate

Connect to your Azure environment to establish a session.

```powershell
Connect-AzAccount

```

Note: Use a managed identity or Key Vault-backed service principal credentials for unattended or scheduled execution.

### 2. Validate with a Dry Run

Perform a dry run to confirm the script functions correctly using synthetic data, ensuring no actual Azure API calls are made.

```powershell
.\AVD-EnvironmentHealth.Generic.ps1 -DryRun -OpenReport -ClientName "XYZ" -ProfileStorageType "ANF" -anfResourceGroupName "XYZ" -UseExistingConnection

```

You can re-run this step using `-ProfileStorageType AzureFiles` or `Both` to preview how the report handles those specific backends.

### 3. Execute Against Live Environment

Run the script against the live environment using your generated configuration file.

```powershell
.\AVD-EnvironmentHealth.Generic.ps1 -ConfigPath .\config.<client>-<env>.json -UseExistingConnection -OpenReport

```

Alternatively, you can supply parameters directly via the command line instead of using a config file. Command-line parameters will override the config file settings.

## Output and Review

* The script execution generates an HTML report.


* Open the generated report and verify that the header correctly displays the Client, Environment, Resource Group, and Profile Storage backend.


* **Confidentiality:** Treat the generated report as Internal/Confidential, as it contains resource names, IDs, and infrastructure health data.


* Distribute the completed report to the responsible AVD operations team for review and follow-up actions.


### Recommended First-Run Checklist

Follow these steps to safely initialize and test the script in a new environment.

**1. Confirm Az modules are current**
Ensure you have the latest Azure PowerShell modules installed.

```powershell
Update-Module Az -Force

```

**2. Authenticate**
Connect to your Azure environment to establish a session.

```powershell
Connect-AzAccount

```

**3. Perform a Dry Run**
Run the script with the `-DryRun` switch. This makes zero Azure calls and confirms that the script parses cleanly.

```powershell
.\AVD-EnvironmentHealth.Generic.ps1 `
    -DryRun -OpenReport -ClientName "Test"

```

**4. Live Run (Non-Production)**
Execute a live run against a non-production subscription first, if available, using your generated configuration file.

```powershell
.\AVD-EnvironmentHealth.Generic.ps1 `
    -ConfigPath .\config.<client>-<env>.json `
    -UseExistingConnection -OpenReport

```

**5. First Production Run**
Execute your first production run and watch the console output live to ensure everything functions as expected.

> **Important:** Do NOT schedule unattended execution until Step 4 has completed cleanly.

```powershell
.\AVD-EnvironmentHealth.Generic.ps1 `
    -ConfigPath .\config.<client>-prod.json `
    -UseExistingConnection -OpenReport

```
