# PsScripts

![PowerShell](https://img.shields.io/badge/Language-PowerShell-blue.svg)
![MIT License](https://img.shields.io/badge/License-MIT-yellow.svg)

A collection of real-world PowerShell scripts created and maintained by **Don Cook**. These tools are used to simplify, automate, and secure IT administration across both on-prem and hybrid Microsoft environments.

---

## 📁 Repository Structure

```
PsScripts/
├── ActiveDirectory/       # On-prem group/user management and reporting
├── AzureAD/               # Cloud mailbox/user automation
├── GeneralAdmin/          # General system and deployment scripts
├── MFA/                   # MFA and identity security scripts
```

---

## 🚀 Featured Scripts

| Script                          | Description |
|---------------------------------|-------------|
| `Audit-AdminAccounts.ps1`       | Reports inactive admin accounts based on naming convention (`a_*`) |
| `Copy-ADUserWithGUI.ps1`        | Interactive GUI for creating users based on a source template |
| `Disable-User.ps1`              | Hybrid account reset + disable with AzureAD + AD |
| `Get-ExpiringADUsers.ps1`       | Identify users with soon-to-expire passwords |
| `Compare-GroupMembers.ps1`      | Compare membership between two AD groups |
| `Audit-MissingAuthenticatorMFA.ps1` | Identify users missing MFA authenticators in Entra ID |

---

## 🧰 Prerequisites

- PowerShell 5.1+ or PowerShell 7+ (where applicable)
- RSAT: Active Directory PowerShell Module
- AzureAD module (installable via `Install-Module AzureAD`)
- Appropriate credentials and administrative rights

---

## 🔒 Usage Considerations

- This repo avoids any real domain references — please **update default domain placeholders** before running.
- Customize output paths (`C:\Reports`, etc.) or make them script parameters for flexibility.

---

## 📜 License

Licensed under the MIT License – feel free to use and adapt the scripts with credit.

---

## 🙌 Author

Don Cook  

[LinkedIn (optional)](https://www.linkedin.com/in/doncook79)  

GitHub: [@reveal79](https://github.com/reveal79)
