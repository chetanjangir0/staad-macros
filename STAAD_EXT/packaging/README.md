# Microsoft Store MSIX

The `Build Microsoft Store MSIX` GitHub Actions workflow creates an unsigned
MSIX for submission to Microsoft Partner Center. The Store signs an accepted
MSIX; the downloaded workflow artifact is a submission package and is not
intended for direct installation.

## Build and submit

1. Create or use a Partner Center developer account.
2. Reserve the `STAAD_EXT` product name.
3. Open the product identity page and copy its **Package/Identity/Name** and
   **Package/Identity/Publisher** values exactly.
4. Run **Actions > Build Microsoft Store MSIX > Run workflow** and supply those
   values plus a four-part version such as `0.1.0.0`.
5. Download the workflow artifact and upload its `.msix` file to the product
   submission in Partner Center.

The package uses the `runFullTrust` restricted capability because the desktop
application must connect to the separately installed STAAD.Pro 2025 OpenSTAAD
COM server. State in the Store submission notes that STAAD.Pro 2025 is required
and must be running with a saved model before using an extension.

## Local package build

Install the Windows 10 or 11 SDK (for `MakeAppx.exe`), build the PyInstaller
folder, and run:

```powershell
pyinstaller --noconfirm --clean STAAD_EXT.spec
.\packaging\Build-Msix.ps1 `
  -IdentityName "PACKAGE_IDENTITY_NAME_FROM_PARTNER_CENTER" `
  -Publisher "CN=PUBLISHER_ID_FROM_PARTNER_CENTER" `
  -Version "0.1.0.0"
```

Do not invent or change the identity values. Partner Center rejects packages
whose manifest identity does not exactly match the reserved product identity.
