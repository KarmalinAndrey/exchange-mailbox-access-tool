# Exchange Online Shared Mailbox Access Tool

Internal PowerShell tool for safely managing access to shared mailboxes
in Microsoft Exchange Online.

---

## Purpose

Granting access to shared mailboxes via Azure Portal is a repetitive
and error-prone task.

This tool automates permission management while keeping production
changes safe and controlled.

---

## Features

- Mandatory **dry-run** mode (no accidental changes)
- Clear permission status (what exists / what is missing)
- Interactive confirmation before applying changes
- Idempotent logic (safe to run multiple times)
- Uses **official Exchange Online PowerShell module**
- No credentials stored in code

---

## How it works

The tool performs the following steps:

1. Validates that the mailbox exists
2. Validates that the user exists
3. Checks existing permissions:
   - FullAccess (Read and manage)
   - SendAs
4. Shows current permission state
5. Requires explicit confirmation before making changes

---

## Usage

---

## Pre-run checklist

Before running the script, make sure that:

1. You are connected to Exchange Online PowerShell:
   ```powershell
   Connect-ExchangeOnline


### Dry-run (required step)

```powershell
.\Add-SharedMailboxAccess.ps1 `
  -Mailbox sales@company.com `
  -User user@company.com `
  -DryRun