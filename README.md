# Windows SysAdmin Scripts (2004-2006)

Batch and VBScript automation scripts developed during my time at the IT department
of a multinational automotive company (Mollet del Vallès, Barcelona, Spain).

## Background

In 2004, I was hired along with a colleague as temporary staff to migrate all computers
at the company's headquarters. The initial task involved going computer by computer,
manually installing software, configuring network drives, and applying security patches.

Seeing how repetitive these tasks were, I developed these scripts to automate the entire
process using `psexec` and VBS/BAT scripts. This allowed us to:

- **Drastically reduce migration time** from weeks to days
- **Mass-deploy security patches** (critical during the Blaster/Sasser worm outbreaks)
- **Automatically configure network drives** based on Active Directory group membership
- **Standardize configuration** across hundreds of workstations

As a result, we finished the project well ahead of schedule, and both of us were
**hired permanently in the IT department**. After this project, the company adopted
more sophisticated deployment systems (Microsoft SMS) for hotfix management.

## Security Notice

> **All hostnames and IP addresses in these scripts have been modified** for security
> purposes. The original infrastructure details are not reflected in this code.

## Technologies

- **VBScript** - Windows automation, WMI, ADSI, COM objects
- **Batch (BAT/CMD)** - Command-line scripting
- **PsExec** (Sysinternals) - Remote execution
- **Active Directory** - LDAP queries for users and groups
- **WMI** - Windows Management Instrumentation

## Folder Structure

| Folder | Description |
|--------|-------------|
| `deployment/` | Mass installation of patches and service packs |
| `network/` | Network drive mapping and hosts file configuration |
| `hardening/` | Security scripts (NetBIOS, UPnP, admin shares) |
| `monitoring/` | File integrity monitoring and printer counters |
| `corporate/` | Corporate configuration (wallpaper, system info) |
| `utilities/` | Miscellaneous utilities |
| `misc/` | Experimental scripts and tests |

## Featured Scripts

### `deployment/massive.bat`
Mass patcher that scans an IP range and automatically deploys security patches
using PsExec. Supports multiple operating systems and languages (8 languages).

```
Usage: massive.bat <IPstart> <IPend> <patch.exe> <reboot:0|1>
Example: massive.bat 192.168.1.1 192.168.1.254 KB824146.exe 0
```

### `network/UnidadesRed.vbs`
Detects the user's Active Directory group membership and automatically maps
the corresponding network drives (Personal, Department, Site) based on the
organizational unit.

### `monitoring/stockmon.vbs`
File integrity monitor. Calculates MD5 hash, detects changes, and sends
email notifications using raw SMTP via netcat. Runs as a background daemon.

### `hardening/disable-netbios.bat`
Disables NetBIOS over TCP/IP across an entire IP range to improve network
security and reduce attack surface.

### `monitoring/xerox-counters.bat`
Collects print counters from multiple Xerox WorkCentre printers via HTTP
scraping and sends a consolidated report via email.

## Historical Context

These scripts were developed during the era of:

- **MS03-039** (DCOM RPC vulnerability) - Exploited by the **Blaster worm** (August 2003)
- **MS04-011** (LSASS vulnerability) - Exploited by the **Sasser worm** (April 2004)
- Mass migration from **Windows 2000 to Windows XP**
- Pre-PowerShell era (PowerShell 1.0 wasn't released until 2006)

The urgency of deploying security patches during the Blaster/Sasser outbreaks was
actually what prompted the company to hire temporary IT staff (including myself)
to accelerate the migration and patching process across all workstations.

## Requirements

- Windows 2000 / Windows XP environment
- Domain Administrator privileges
- PsExec and other Sysinternals tools
- Network access to target machines (admin shares)

## Notes

- These are **legacy scripts** preserved for educational and portfolio purposes
- The code is in its original 2004-2006 state
- Some scripts contain hardcoded paths that would need adjustment for other environments
- Comments are primarily in Spanish (the original development language)

## Author

**Jordi Corrales**
IT Systems Department (2004-2006)

---

*These scripts are shared for educational purposes and as a historical portfolio piece
demonstrating Windows administration automation in the pre-PowerShell era.*
